# -*- coding: utf-8 -*-
"""
AI任务中台 - 推广支持Bot
功能：知识库问答 + 需求收集 + Excel汇总导出 + 管理后台
"""
import os
import json
import time
import re
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Render 等平台用 /tmp 作为可写目录；本地用 data 目录
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(BASE_DIR, "data"))
REQUESTS_FILE = os.path.join(DATA_DIR, "requests.json")
KB_FILE = os.path.join(BASE_DIR, "knowledge_base.json")

os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__)

# ============================================================
# 知识库加载
# ============================================================
def load_kb():
    with open(KB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

KB = load_kb()

def reload_kb():
    """重新加载知识库（修改后生效）"""
    global KB
    KB = load_kb()

# ============================================================
# 问答匹配引擎
# ============================================================
def normalize(text):
    """简单归一化：转小写、去标点空格"""
    text = text.lower().strip()
    text = re.sub(r"[\s，。、！？,.!?~～\-_]", "", text)
    return text

def score_match(user_input, faq_item):
    """
    关键词匹配打分：
    - 每命中一个关键词加分
    - 命中question本身加分更高
    - 长关键词权重更高
    """
    score = 0
    norm_input = normalize(user_input)
    if not norm_input:
        return 0

    # 关键词命中
    for kw in faq_item.get("keywords", []):
        kw_n = normalize(kw)
        if not kw_n:
            continue
        if kw_n in norm_input:
            # 长关键词权重高
            score += 1 + len(kw_n) * 0.3

    # question 整体命中
    q_n = normalize(faq_item.get("question", ""))
    if q_n and q_n in norm_input:
        score += 5

    # question 分词命中（按字滑动）
    if q_n:
        overlap = 0
        for ch in q_n:
            if ch in norm_input:
                overlap += 1
        ratio = overlap / max(len(q_n), 1)
        score += ratio * 2

    return score

def find_answer(user_input):
    """返回最佳匹配的答案，若无匹配返回None"""
    kb = load_kb()
    best_score = 0
    best_item = None
    for item in kb["faq"]:
        s = score_match(user_input, item)
        if s > best_score:
            best_score = s
            best_item = item

    # 阈值：经验值，低于此认为没匹配上
    if best_score < 1.5 or best_item is None:
        return None, None, 0
    return best_item["answer"], best_item, best_score

# ============================================================
# 需求记录
# ============================================================
def load_requests():
    if os.path.exists(REQUESTS_FILE):
        try:
            with open(REQUESTS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_requests(data):
    with open(REQUESTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def add_request(org, contact, rtype, desc, extra=""):
    """新增一条需求记录"""
    reqs = load_requests()
    record = {
        "id": str(int(time.time() * 1000)),
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "org": org,
        "contact": contact,
        "type": rtype,
        "description": desc,
        "status": "待处理",
        "extra": extra
    }
    reqs.append(record)
    save_requests(reqs)
    return record

# ============================================================
# 路由 - 页面
# ============================================================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/admin")
def admin():
    return render_template("admin.html")

# ============================================================
# API - 问答
# ============================================================
@app.route("/api/chat", methods=["POST"])
def api_chat():
    data = request.get_json(silent=True) or {}
    user_msg = data.get("message", "").strip()
    if not user_msg:
        return jsonify({"ok": False, "error": "消息不能为空"}), 400

    answer, item, score = find_answer(user_msg)

    if answer:
        return jsonify({
            "ok": True,
            "matched": True,
            "answer": answer,
            "category": item.get("category", ""),
            "question": item.get("question", ""),
            "score": round(score, 2)
        })
    else:
        # 未匹配，引导提需求
        return jsonify({
            "ok": True,
            "matched": False,
            "answer": "抱歉，这个问题我暂时还没有现成的答案。\n\n您可以通过下方的「我要提需求」按钮，把问题或需求留给我们，项目经理赖雨晴会尽快跟进。常见问题也可以点击下方快捷提问试试。",
            "suggest_feedback": True
        })

# 快捷问题列表
@app.route("/api/quick_questions", methods=["GET"])
def api_quick_questions():
    kb = load_kb()
    items = [{"id": f["id"], "question": f["question"]} for f in kb["faq"]]
    return jsonify({"ok": True, "questions": items})

# ============================================================
# API - 需求收集
# ============================================================
@app.route("/api/feedback", methods=["POST"])
def api_feedback():
    data = request.get_json(silent=True) or {}
    org = data.get("org", "").strip()
    contact = data.get("contact", "").strip()
    rtype = data.get("type", "咨询").strip()
    desc = data.get("description", "").strip()
    extra = data.get("extra", "").strip()

    if not desc:
        return jsonify({"ok": False, "error": "问题描述不能为空"}), 400

    record = add_request(org, contact, rtype, desc, extra)
    return jsonify({
        "ok": True,
        "message": "已收到您的反馈，项目经理赖雨晴会尽快跟进，感谢支持！",
        "record_id": record["id"]
    })

# ============================================================
# API - 管理后台
# ============================================================
@app.route("/api/requests", methods=["GET"])
def api_requests():
    reqs = load_requests()
    # 支持筛选
    org = request.args.get("org", "").strip()
    rtype = request.args.get("type", "").strip()
    status = request.args.get("status", "").strip()
    result = reqs
    if org:
        result = [r for r in result if org in r.get("org", "")]
    if rtype:
        result = [r for r in result if r.get("type") == rtype]
    if status:
        result = [r for r in result if r.get("status") == status]
    return jsonify({"ok": True, "total": len(result), "data": result})

@app.route("/api/requests/<rid>/status", methods=["POST"])
def api_update_status(rid):
    data = request.get_json(silent=True) or {}
    new_status = data.get("status", "").strip()
    reqs = load_requests()
    for r in reqs:
        if r["id"] == rid:
            r["status"] = new_status
            save_requests(reqs)
            return jsonify({"ok": True, "message": "状态已更新"})
    return jsonify({"ok": False, "error": "未找到记录"}), 404

@app.route("/api/requests/<rid>", methods=["DELETE"])
def api_delete_request(rid):
    reqs = load_requests()
    new_reqs = [r for r in reqs if r["id"] != rid]
    if len(new_reqs) == len(reqs):
        return jsonify({"ok": False, "error": "未找到记录"}), 404
    save_requests(new_reqs)
    return jsonify({"ok": True, "message": "已删除"})

# 统计
@app.route("/api/stats", methods=["GET"])
def api_stats():
    reqs = load_requests()
    by_type = {}
    by_status = {}
    by_org = {}
    for r in reqs:
        t = r.get("type", "未分类")
        s = r.get("status", "待处理")
        o = r.get("org", "未填写") or "未填写"
        by_type[t] = by_type.get(t, 0) + 1
        by_status[s] = by_status.get(s, 0) + 1
        by_org[o] = by_org.get(o, 0) + 1
    return jsonify({
        "ok": True,
        "total": len(reqs),
        "by_type": by_type,
        "by_status": by_status,
        "by_org": by_org
    })

# Excel导出
@app.route("/api/export_excel", methods=["GET"])
def api_export_excel():
    reqs = load_requests()
    wb = Workbook()
    ws = wb.active
    ws.title = "需求汇总"

    # 表头
    headers = ["序号", "提交时间", "机构名称", "联系人", "类型", "问题描述", "补充说明", "状态"]
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 数据
    for i, r in enumerate(reqs, 1):
        row = i + 1
        values = [
            i,
            r.get("time", ""),
            r.get("org", ""),
            r.get("contact", ""),
            r.get("type", ""),
            r.get("description", ""),
            r.get("extra", ""),
            r.get("status", "")
        ]
        for col, v in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=v if v else "")
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = thin_border

    # 列���
    widths = [6, 20, 18, 14, 10, 50, 30, 10]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + col) if col <= 26 else "A"].width = w

    # 冻结首行
    ws.freeze_panes = "A2"

    # 保存到临时文件
    out_path = os.path.join(DATA_DIR, "需求汇总.xlsx")
    wb.save(out_path)

    filename = f"AI任务中台_需求汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        out_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ============================================================
# API - 知识库管理
# ============================================================
@app.route("/api/kb", methods=["GET"])
def api_kb_list():
    kb = load_kb()
    return jsonify({"ok": True, "meta": kb.get("meta", {}), "faq": kb["faq"]})

@app.route("/api/kb", methods=["POST"])
def api_kb_add():
    """新增知识库条目"""
    data = request.get_json(silent=True) or {}
    question = data.get("question", "").strip()
    answer = data.get("answer", "").strip()
    keywords_raw = data.get("keywords", "").strip()
    category = data.get("category", "其他").strip()

    if not question or not answer:
        return jsonify({"ok": False, "error": "问题和答案不能为空"}), 400

    # 关键词支持逗号分隔
    keywords = [k.strip() for k in re.split(r"[，,、\n]", keywords_raw) if k.strip()]
    # 自动把问题本身拆成关键词补充
    for ch in question:
        if ch not in keywords and ch not in "？?的了吗呢啊是有什么":
            keywords.append(ch)

    kb = load_kb()
    import uuid
    new_id = "custom_" + uuid.uuid4().hex[:8]
    new_item = {
        "id": new_id,
        "category": category,
        "question": question,
        "keywords": keywords,
        "answer": answer
    }
    kb["faq"].append(new_item)
    with open(KB_FILE, "w", encoding="utf-8") as f:
        json.dump(kb, f, ensure_ascii=False, indent=2)
    reload_kb()
    return jsonify({"ok": True, "message": "已新增知识库条目", "id": new_id})

@app.route("/api/kb/<kid>", methods=["PUT"])
def api_kb_update(kid):
    """修改知识库条目"""
    data = request.get_json(silent=True) or {}
    kb = load_kb()
    for item in kb["faq"]:
        if item["id"] == kid:
            if "question" in data:
                item["question"] = data["question"].strip()
            if "answer" in data:
                item["answer"] = data["answer"].strip()
            if "keywords" in data:
                kws = data["keywords"]
                if isinstance(kws, str):
                    item["keywords"] = [k.strip() for k in re.split(r"[，,、\n]", kws) if k.strip()]
                else:
                    item["keywords"] = kws
            if "category" in data:
                item["category"] = data["category"].strip()
            with open(KB_FILE, "w", encoding="utf-8") as f:
                json.dump(kb, f, ensure_ascii=False, indent=2)
            reload_kb()
            return jsonify({"ok": True, "message": "已修改"})
    return jsonify({"ok": False, "error": "未找到条目"}), 404

@app.route("/api/kb/<kid>", methods=["DELETE"])
def api_kb_delete(kid):
    """删除知识库条目"""
    kb = load_kb()
    new_faq = [f for f in kb["faq"] if f["id"] != kid]
    if len(new_faq) == len(kb["faq"]):
        return jsonify({"ok": False, "error": "未找到条目"}), 404
    kb["faq"] = new_faq
    with open(KB_FILE, "w", encoding="utf-8") as f:
        json.dump(kb, f, ensure_ascii=False, indent=2)
    reload_kb()
    return jsonify({"ok": True, "message": "已删除"})


if __name__ == "__main__":
    # 写入若干示例数据（首次运行）
    if not os.path.exists(REQUESTS_FILE) or len(load_requests()) == 0:
        add_request("示例机构A", "张三", "咨询", "想了解AI质检是怎么判断回访有效性的，准确率怎么样？", "")
        add_request("示例机构B", "李四", "需求", "希望话术能支持按产品线区分，不同产品话术不同。", "")
        add_request("示例机构A", "张三", "bug", "总览页面加载比较慢，有时候要等好几秒。", "")

    print("=" * 50)
    print("AI任务中台 推广支持Bot 已启动")
    print("  使用者入口: http://localhost:5000/")
    print("  管理后台:   http://localhost:5000/admin")
    print("=" * 50)

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
