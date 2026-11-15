/* Address book front-end: no framework, no build step. */

const TYPE_LABEL = {
  phone: "手机",
  email: "邮箱",
  wechat: "微信",
  qq: "QQ",
  address: "地址",
  other: "其他",
};

const state = {
  contacts: [],
  stats: { total: 0, bookmarked: 0, methods: 0 },
  keyword: "",
  onlyFav: false,
  selectedId: null,
};

const $ = (sel) => document.querySelector(sel);
const listEl = $("#contact-list");
const detailEl = $("#detail");
const editor = $("#editor");
const form = $("#editor-form");
const rowsEl = $("#method-rows");

/* ------------------------------ helpers ------------------------------ */

function toast(message, isError = false) {
  const el = $("#toast");
  el.textContent = message;
  el.className = "toast show" + (isError ? " error" : "");
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { el.className = "toast"; }, 2600);
}

function initial(name) {
  const trimmed = (name || "").trim();
  return trimmed ? trimmed.slice(-2) : "?";
}

function avatarColor(name) {
  let hash = 0;
  for (const ch of name || "?") hash = (hash * 31 + ch.codePointAt(0)) % 360;
  return `hsl(${hash}, 62%, 52%)`;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[ch]));
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  let payload = {};
  try { payload = await response.json(); } catch { /* ignore */ }
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.error || `请求失败 (${response.status})`);
  }
  return payload;
}

/* ------------------------------ rendering ------------------------------ */

function renderStats() {
  $("#stat-total").textContent = state.stats.total;
  $("#stat-fav").textContent = state.stats.bookmarked;
  $("#stat-methods").textContent = state.stats.methods;
}

function renderList() {
  listEl.innerHTML = "";
  state.contacts.forEach((contact) => {
    const li = document.createElement("li");
    li.dataset.id = contact.id;
    if (contact.id === state.selectedId) li.classList.add("active");

    const first = contact.methods && contact.methods.length ? contact.methods[0].value : "暂无联系方式";
    li.innerHTML = `
      <div class="avatar" style="background:${avatarColor(contact.name)}">${escapeHtml(initial(contact.name))}</div>
      <div class="who">
        <b>${escapeHtml(contact.name)}</b>
        <small>${escapeHtml(first)}</small>
      </div>
      ${contact.bookmarked ? '<span class="fav-mark" title="已收藏">★</span>' : ""}
    `;
    li.addEventListener("click", () => selectContact(contact.id));
    listEl.appendChild(li);
  });
  $("#empty-hint").style.display = state.contacts.length ? "none" : "block";
  if (!state.contacts.length) {
    $("#empty-hint").textContent = state.keyword || state.onlyFav
      ? "没有匹配的联系人。"
      : "还没有联系人，点右上角“新建联系人”开始吧。";
  }
}

function renderDetail() {
  const contact = state.contacts.find((c) => c.id === state.selectedId);
  if (!contact) {
    detailEl.innerHTML = `
      <div class="detail-empty">
        <div class="big">👤</div>
        <h2>选择左侧联系人查看详情</h2>
        <p>支持收藏标记、多种联系方式、Excel 导入导出。</p>
      </div>`;
    return;
  }

  const methods = (contact.methods || []).map((method) => `
    <div class="method">
      <span class="kind">${escapeHtml(TYPE_LABEL[method.type] || "其他")}</span>
      <span class="value">${escapeHtml(method.value)}</span>
      <button class="btn ghost small copy" data-copy="${escapeHtml(method.value)}">复制</button>
    </div>`).join("");

  detailEl.innerHTML = `
    <div class="detail-head">
      <div class="avatar" style="background:${avatarColor(contact.name)}">${escapeHtml(initial(contact.name))}</div>
      <div>
        <h2>${escapeHtml(contact.name)}</h2>
        <div class="meta">
          ${contact.bookmarked ? '<span class="tag">★ 已收藏</span>' : ""}
          更新于 ${escapeHtml(contact.updatedAt || "-")}
        </div>
      </div>
      <div class="spacer"></div>
      <button class="btn ghost" id="btn-toggle-fav">${contact.bookmarked ? "取消收藏" : "★ 收藏"}</button>
    </div>

    <div class="methods">
      ${methods || '<p style="color:#6b7a99">该联系人还没有联系方式。</p>'}
    </div>

    ${contact.note ? `<div class="note">${escapeHtml(contact.note)}</div>` : ""}

    <div class="detail-foot">
      <button class="btn primary" id="btn-edit">编辑</button>
      <button class="btn ghost" id="btn-delete" style="color:#d64545">删除</button>
    </div>
  `;

  $("#btn-toggle-fav").addEventListener("click", () => toggleBookmark(contact.id));
  $("#btn-edit").addEventListener("click", () => openEditor(contact));
  $("#btn-delete").addEventListener("click", () => deleteContact(contact));
  detailEl.querySelectorAll("[data-copy]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(btn.dataset.copy);
        toast("已复制到剪贴板");
      } catch {
        toast("复制失败，请手动选择文本", true);
      }
    });
  });
}

/* ------------------------------ data ops ------------------------------ */

async function refresh(keepSelection = true) {
  const params = new URLSearchParams();
  if (state.keyword) params.set("q", state.keyword);
  if (state.onlyFav) params.set("bookmarked", "1");
  const payload = await api(`/api/contacts?${params.toString()}`);
  state.contacts = payload.contacts;
  state.stats = payload.stats;
  if (keepSelection && state.selectedId && !state.contacts.some((c) => c.id === state.selectedId)) {
    state.selectedId = null;
  }
  renderStats();
  renderList();
  renderDetail();
}

function selectContact(id) {
  state.selectedId = id;
  if (id) {
    history.replaceState(null, "", "#" + id);
  } else {
    history.replaceState(null, "", location.pathname);
  }
  renderList();
  renderDetail();
}

async function toggleBookmark(id) {
  try {
    const payload = await api(`/api/contacts/${id}/toggle-bookmark`, { method: "POST" });
    toast(payload.contact.bookmarked ? "已加入收藏 ★" : "已取消收藏");
    await refresh();
  } catch (error) {
    toast(error.message, true);
  }
}

async function deleteContact(contact) {
  if (!confirm(`确定删除联系人“${contact.name}”吗？`)) return;
  try {
    await api(`/api/contacts/${contact.id}`, { method: "DELETE" });
    state.selectedId = null;
    toast("已删除");
    await refresh();
  } catch (error) {
    toast(error.message, true);
  }
}

/* ------------------------------ editor ------------------------------ */

function addMethodRow(type = "phone", value = "") {
  const fragment = $("#method-row-template").content.cloneNode(true);
  const row = fragment.querySelector(".method-row");
  row.querySelector(".m-type").value = type;
  row.querySelector(".m-value").value = value;
  row.querySelector(".icon-btn").addEventListener("click", () => row.remove());
  rowsEl.appendChild(row);
  return row;
}

function openEditor(contact = null) {
  form.dataset.id = contact ? contact.id : "";
  $("#editor-title").textContent = contact ? "编辑联系人" : "新建联系人";
  $("#f-name").value = contact ? contact.name : "";
  $("#f-fav").checked = contact ? !!contact.bookmarked : false;
  $("#f-note").value = contact ? contact.note || "" : "";
  rowsEl.innerHTML = "";
  if (contact && contact.methods.length) {
    contact.methods.forEach((m) => addMethodRow(m.type, m.value));
  } else {
    addMethodRow("phone", "");
  }
  editor.showModal();
  $("#f-name").focus();
}

function collectForm() {
  const methods = [];
  rowsEl.querySelectorAll(".method-row").forEach((row) => {
    const value = row.querySelector(".m-value").value.trim();
    if (value) methods.push({ type: row.querySelector(".m-type").value, value });
  });
  return {
    name: $("#f-name").value.trim(),
    bookmarked: $("#f-fav").checked,
    note: $("#f-note").value.trim(),
    methods,
  };
}

async function saveContact(event) {
  event.preventDefault();
  const payload = collectForm();
  if (!payload.name) {
    toast("请填写姓名", true);
    $("#f-name").focus();
    return;
  }
  const id = form.dataset.id;
  try {
    if (id) {
      await api(`/api/contacts/${id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      state.selectedId = id;
      toast("已保存修改");
    } else {
      const result = await api("/api/contacts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      state.selectedId = result.contact.id;
      toast("已添加联系人");
    }
    editor.close();
    await refresh();
  } catch (error) {
    toast(error.message, true);
  }
}

/* ------------------------------ import / export ------------------------------ */

async function importExcel(file) {
  if (!file) return;
  if (!/\.xlsx$/i.test(file.name)) {
    toast("请选择 .xlsx 文件", true);
    return;
  }
  try {
    const payload = await api("/api/import", {
      method: "POST",
      headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      body: file,
    });
    const { imported, updated, methodsAdded } = payload.stats;
    toast(`导入完成：新增 ${imported} 人，更新 ${updated} 人，新增 ${methodsAdded} 条联系方式`);
    await refresh();
  } catch (error) {
    toast(error.message, true);
  }
}

/* ------------------------------ wiring ------------------------------ */

$("#btn-new").addEventListener("click", () => openEditor(null));
$("#btn-add-method").addEventListener("click", () => addMethodRow());
$("#btn-export").addEventListener("click", () => { window.location.href = "/api/export"; });

$("#btn-fav").addEventListener("click", (event) => {
  state.onlyFav = !state.onlyFav;
  event.currentTarget.classList.toggle("active", state.onlyFav);
  refresh();
});

let searchTimer = null;
$("#search").addEventListener("input", (event) => {
  clearTimeout(searchTimer);
  const value = event.target.value;
  searchTimer = setTimeout(() => {
    state.keyword = value.trim();
    refresh();
  }, 180);
});

$("#file-input").addEventListener("change", (event) => {
  importExcel(event.target.files[0]);
  event.target.value = "";
});

form.addEventListener("submit", saveContact);
form.querySelectorAll("[data-close]").forEach((btn) => btn.addEventListener("click", () => editor.close()));
editor.addEventListener("click", (event) => { if (event.target === editor) editor.close(); });

refresh()
  .then(() => {
    // deep link: /#<contact id> opens that contact directly
    const hashId = location.hash.replace("#", "");
    if (hashId === "new") {
      openEditor(null);
    } else if (hashId && state.contacts.some((c) => c.id === hashId)) {
      selectContact(hashId);
    } else if (state.contacts.length) {
      selectContact(state.contacts[0].id);
    }
  })
  .catch((error) => toast(error.message, true));
