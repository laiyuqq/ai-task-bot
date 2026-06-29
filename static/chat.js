// ===== AI任务中台 推广支持Bot 前端逻辑 =====
const messagesEl = document.getElementById("messages");
const inputEl = document.getElementById("input");
const sendBtn = document.getElementById("sendBtn");
const quickBar = document.getElementById("quickBar");
const feedbackModal = document.getElementById("feedbackModal");

// 对话历史（发给DeepSeek实现多轮对话）
let chatHistory = [];

// 追加消息
function addMessage(text, sender) {
  const msg = document.createElement("div");
  msg.className = "msg " + sender;
  const avatar = document.createElement("div");
  avatar.className = "avatar " + sender + "-avatar";
  avatar.textContent = sender === "bot" ? "AI" : "我";
  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.innerHTML = text;
  msg.appendChild(avatar);
  msg.appendChild(bubble);
  messagesEl.appendChild(msg);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return bubble;
}

// 打字动画
function showTyping() {
  const msg = document.createElement("div");
  msg.className = "msg bot";
  msg.id = "typingMsg";
  msg.innerHTML = '<div class="avatar bot-avatar">AI</div><div class="bubble"><div class="typing"><span></span><span></span><span></span></div></div>';
  messagesEl.appendChild(msg);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}
function hideTyping() {
  const t = document.getElementById("typingMsg");
  if (t) t.remove();
}

// 发送消息
async function send(text) {
  if (!text.trim()) return;
  addMessage(escapeHtml(text), "user");
  inputEl.value = "";
  sendBtn.disabled = true;
  showTyping();

  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: text, history: chatHistory })
    });
    const data = await res.json();
    hideTyping();

    if (data.ok) {
      // 简单的markdown渲染：加粗、换行、列表
      const html = formatAnswer(data.answer);
      addMessage(html, "bot");

      // 记录历史
      chatHistory.push({ role: "user", content: text });
      chatHistory.push({ role: "assistant", content: data.answer });
      // 保留最近10条
      if (chatHistory.length > 20) chatHistory = chatHistory.slice(-20);

      if (data.suggest_feedback) {
        addMessage('点击下方「📝 我要提需求」按钮，把您的需求留给我们，项目经理赖雨晴会尽快跟进 👇', "bot");
      }
    } else {
      addMessage("出了点小问题，请稍后重试。", "bot");
    }
  } catch (e) {
    hideTyping();
    addMessage("网络异常，请稍后重试。", "bot");
  }
  sendBtn.disabled = false;
  inputEl.focus();
}

// 简单markdown格式化
function formatAnswer(text) {
  let s = escapeHtml(text);
  // 加粗 **text**
  s = s.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  // 换行
  s = s.replace(/\n/g, "<br>");
  // 列表项 • 或 - 开头
  s = s.replace(/^<br>[•\-]\s*/gm, "<br>• ");
  return s;
}

// HTML转义
function escapeHtml(s) {
  if (!s) return "";
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// 事件绑定
sendBtn.addEventListener("click", () => send(inputEl.value));
inputEl.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    send(inputEl.value);
  }
});

// 快捷问题
async function loadQuickQuestions() {
  try {
    const res = await fetch("/api/quick_questions");
    const data = await res.json();
    if (data.ok) {
      data.questions.forEach(q => {
        const chip = document.createElement("div");
        chip.className = "quick-chip";
        chip.textContent = q.question;
        chip.addEventListener("click", () => send(q.question));
        quickBar.appendChild(chip);
      });
    }
  } catch (e) { console.error(e); }
}
loadQuickQuestions();

// ===== 反馈表单 =====
const btnFeedback = document.getElementById("btnFeedback");
const modalClose = document.getElementById("modalClose");
const btnCancel = document.getElementById("btnCancel");
const btnSubmit = document.getElementById("btnSubmit");

function openModal() {
  feedbackModal.style.display = "flex";
  document.getElementById("fOrg").focus();
}
function closeModal() {
  feedbackModal.style.display = "none";
  document.getElementById("formTip").textContent = "";
  document.getElementById("formTip").className = "form-tip";
}

btnFeedback.addEventListener("click", openModal);
modalClose.addEventListener("click", closeModal);
btnCancel.addEventListener("click", closeModal);
feedbackModal.addEventListener("click", (e) => {
  if (e.target === feedbackModal) closeModal();
});

// 提交反馈
btnSubmit.addEventListener("click", async () => {
  const org = document.getElementById("fOrg").value.trim();
  const contact = document.getElementById("fContact").value.trim();
  const type = document.getElementById("fType").value;
  const desc = document.getElementById("fDesc").value.trim();
  const extra = document.getElementById("fExtra").value.trim();
  const tip = document.getElementById("formTip");

  if (!org) { tip.textContent = "请填写机构名称"; return; }
  if (!contact) { tip.textContent = "请填写联系人"; return; }
  if (!desc) { tip.textContent = "请填写详细描述"; return; }

  btnSubmit.disabled = true;
  btnSubmit.textContent = "提交中…";

  try {
    const res = await fetch("/api/feedback", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ org, contact, type, description: desc, extra })
    });
    const data = await res.json();
    if (data.ok) {
      tip.className = "form-tip success";
      tip.textContent = "✅ " + data.message;
      document.getElementById("fOrg").value = "";
      document.getElementById("fContact").value = "";
      document.getElementById("fDesc").value = "";
      document.getElementById("fExtra").value = "";
      setTimeout(() => {
        closeModal();
        addMessage("✅ 您的需求已提交成功！项目经理赖雨晴会尽快跟进，感谢您的反馈。", "bot");
      }, 1200);
    } else {
      tip.textContent = "提交失败：" + (data.error || "未知错误");
    }
  } catch (e) {
    tip.textContent = "网络异常，请稍后重试";
  }
  btnSubmit.disabled = false;
  btnSubmit.textContent = "提交";
});
