const form = document.querySelector("#chatForm");
const input = document.querySelector("#messageInput");
const messages = document.querySelector("#messages");
const profileList = document.querySelector("#profileList");
const tagList = document.querySelector("#tagList");
const recommendationList = document.querySelector("#recommendationList");
const phaseBadge = document.querySelector("#phaseBadge");
const demoButton = document.querySelector("#demoButton");

const phaseText = {
  question: "追问信息",
  recommend: "已推荐",
  red_flag: "需就医评估",
};

function addMessage(role, text) {
  const node = document.createElement("div");
  node.className = `message ${role}`;
  node.textContent = text;
  messages.appendChild(node);
  messages.scrollTop = messages.scrollHeight;
}

function renderProfile(profile) {
  profileList.innerHTML = "";
  if (!profile.length) {
    profileList.className = "profile-list empty";
    profileList.textContent = "暂无画像信息";
    return;
  }
  profileList.className = "profile-list";
  profile.forEach((item) => {
    const node = document.createElement("article");
    node.className = "profile-item";
    const value = Array.isArray(item.value) ? item.value.join("、") : item.value;
    node.innerHTML = `<strong>${item.label}</strong><small>${value}</small><small>依据：${item.evidence} · 置信度 ${Math.round(
      item.confidence * 100,
    )}%</small>`;
    profileList.appendChild(node);
  });
}

function renderTags(tags) {
  tagList.innerHTML = "";
  if (!tags.length) {
    tagList.className = "tag-list empty";
    tagList.textContent = "暂无标签";
    return;
  }
  tagList.className = "tag-list";
  tags.forEach((tag) => {
    const node = document.createElement("span");
    node.className = "tag";
    node.textContent = tag;
    tagList.appendChild(node);
  });
}

function renderRecommendations(items) {
  recommendationList.innerHTML = "";
  if (!items.length) {
    recommendationList.className = "recommendation-list empty";
    recommendationList.textContent = "信息足够后生成排序";
    return;
  }
  recommendationList.className = "recommendation-list";
  items.forEach((item, index) => {
    const node = document.createElement("article");
    node.className = "recommendation-item";
    node.innerHTML = `<strong>${index + 1}. ${item.intervention}</strong><small>评分 ${item.score} · ${
      item.category
    }</small><small>${item.summary}</small>`;
    recommendationList.appendChild(node);
  });
}

async function sendMessage(text) {
  addMessage("user", text);
  input.value = "";
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: "default", message: text }),
  });
  const data = await response.json();
  if (data.error) {
    addMessage("assistant", `处理失败：${data.error}`);
    return;
  }
  addMessage("assistant", data.reply);
  phaseBadge.textContent = phaseText[data.phase] || "处理中";
  renderProfile(data.profile || []);
  renderTags(data.tags || []);
  renderRecommendations(data.recommendations || []);
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const text = input.value.trim();
  if (!text) return;
  sendMessage(text);
});

demoButton.addEventListener("click", () => {
  input.value = "我爸 72 岁，最近记性差一点，平时不太爱运动，喜欢跟邻居打麻将。";
  input.focus();
});

addMessage(
  "assistant",
  "请先描述对象的年龄、记忆变化、慢病、运动限制、睡眠、社交和听力视力情况。我会先补齐画像，再根据 Excel 证据库排序建议。",
);
