/* === Curry Stock - フロント JS === */

// ▼▼▼ ここを自分の Apps Script Web App URL に書き換える ▼▼▼
const API_URL = 'https://script.google.com/macros/s/AKfycbwakrB-QucMZKmY1u736DtS2OISpyVnllPDzHsL_mNwItRzYDqkBRy3MIAIAOANrpNy0A/exec';
// ▲▲▲ デプロイ後に取得する URL を貼り付ける ▲▲▲

// === 状態 ===
const state = {
  items: [],
  categories: [],
  config: {},
  activeCategory: 'all',
  modalCode: null,
};

// === API クライアント ===
async function apiGet(action, params = {}) {
  const url = new URL(API_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'API error');
  return json.data;
}

async function apiPost(payload) {
  // GAS の Web App は CORS のため text/plain で送るのが定番(プリフライトを避ける)
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'API error');
  return json.data;
}

// === 初期化 ===
document.addEventListener('DOMContentLoaded', () => {
  bindTabs();
  bindModal();
  bindAddForm();
  document.getElementById('reloadHistory').addEventListener('click', loadHistory);
  loadStock();
});

// === タブ切り替え ===
function bindTabs() {
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      const target = btn.dataset.tab;
      document.getElementById('tab-' + target).classList.add('active');
      if (target === 'history') loadHistory();
      if (target === 'add') populateAddCategories();
    });
  });
}

// === 在庫一覧読み込み ===
async function loadStock() {
  try {
    const data = await apiGet('list');
    state.items = data.items;
    state.categories = data.categories;
    state.config = data.config || {};
    renderCategoryTabs();
    renderItems();
    renderAlertBar();
  } catch (err) {
    document.getElementById('itemsList').innerHTML =
      `<div class="empty">読み込み失敗: ${escapeHtml(err.message)}<br><br>API_URL を app.js で正しく設定したか確認してください。</div>`;
  }
}

function renderCategoryTabs() {
  const wrap = document.getElementById('categoryTabs');
  const cats = [{ name: 'all', icon: '🍛', label: 'すべて' }]
    .concat(state.categories.map(c => ({ name: c.name, icon: c.icon, label: c.name })));
  wrap.innerHTML = cats.map(c =>
    `<button class="cat-pill ${state.activeCategory === c.name ? 'active' : ''}" data-cat="${escapeAttr(c.name)}">
       ${c.icon || ''} ${escapeHtml(c.label)}
     </button>`
  ).join('');
  wrap.querySelectorAll('.cat-pill').forEach(p => {
    p.addEventListener('click', () => {
      state.activeCategory = p.dataset.cat;
      renderCategoryTabs();
      renderItems();
    });
  });
}

function renderItems() {
  const list = document.getElementById('itemsList');
  let items = state.items;
  if (state.activeCategory !== 'all') {
    items = items.filter(i => i.category === state.activeCategory);
  }
  if (items.length === 0) {
    list.innerHTML = '<div class="empty">商品がありません</div>';
    return;
  }
  list.innerHTML = items.map(it => {
    const isLow = it.minStock > 0 && it.stock <= it.minStock;
    return `
      <div class="item-card ${isLow ? 'low' : ''}" data-code="${escapeAttr(it.code)}">
        <div class="item-info">
          <div class="item-name">${escapeHtml(it.name)}</div>
          <div class="item-meta">${escapeHtml(it.code)} · ${escapeHtml(it.category)} · 最低 ${it.minStock}${escapeHtml(it.unit)}</div>
        </div>
        <div class="item-stock">
          <span class="stock-num">${formatNum(it.stock)}</span>
          <span class="stock-unit">${escapeHtml(it.unit)}</span>
        </div>
      </div>`;
  }).join('');
  list.querySelectorAll('.item-card').forEach(card => {
    card.addEventListener('click', () => openModal(card.dataset.code));
  });
}

function renderAlertBar() {
  const bar = document.getElementById('alertBar');
  const lows = state.items.filter(i => i.minStock > 0 && i.stock <= i.minStock);
  if (lows.length === 0) {
    bar.classList.add('hidden');
    return;
  }
  bar.classList.remove('hidden');
  bar.textContent = `在庫が少ない商品が ${lows.length} 件あります: ${lows.slice(0, 3).map(i => i.name).join('、')}${lows.length > 3 ? ' ほか' : ''}`;
}

// === モーダル ===
function bindModal() {
  document.getElementById('modalClose').addEventListener('click', closeModal);
  document.getElementById('modal').addEventListener('click', (e) => {
    if (e.target.id === 'modal') closeModal();
  });
  document.querySelectorAll('.qty-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const input = document.getElementById('modalQty');
      const cur = Number(input.value) || 0;
      const delta = Number(btn.dataset.delta);
      input.value = Math.max(0, cur + delta);
    });
  });
  document.querySelectorAll('.preset').forEach(btn => {
    btn.addEventListener('click', () => {
      document.getElementById('modalQty').value = btn.dataset.val;
    });
  });
  document.getElementById('btnIn').addEventListener('click', () => submitInOut('入庫'));
  document.getElementById('btnOut').addEventListener('click', () => submitInOut('出庫'));
}

function openModal(code) {
  const item = state.items.find(i => i.code === code);
  if (!item) return;
  state.modalCode = code;
  document.getElementById('modalItemName').textContent = item.name;
  document.getElementById('modalCurrentStock').textContent = formatNum(item.stock);
  document.getElementById('modalUnit').textContent = item.unit;
  document.getElementById('modalQty').value = 1;
  document.getElementById('modalNote').value = '';
  document.getElementById('modal').classList.remove('hidden');
}
function closeModal() {
  document.getElementById('modal').classList.add('hidden');
  state.modalCode = null;
}

async function submitInOut(type) {
  const qty = Number(document.getElementById('modalQty').value);
  if (!qty || qty <= 0) {
    toast('数量を入力してください', true);
    return;
  }
  const note = document.getElementById('modalNote').value.trim();
  const operator = document.getElementById('operator').value;
  const code = state.modalCode;

  const btnIn = document.getElementById('btnIn');
  const btnOut = document.getElementById('btnOut');
  btnIn.disabled = btnOut.disabled = true;

  try {
    const result = await apiPost({
      action: 'inOut',
      code, type, qty, operator, note
    });
    toast(`${type} OK: ${result.name} は ${formatNum(result.stockAfter)} に`);
    closeModal();
    await loadStock();
  } catch (err) {
    toast('失敗: ' + err.message, true);
  } finally {
    btnIn.disabled = btnOut.disabled = false;
  }
}

// === 履歴 ===
async function loadHistory() {
  const list = document.getElementById('historyList');
  list.innerHTML = '<div class="loading">読み込み中…</div>';
  try {
    const rows = await apiGet('history', { limit: 100 });
    if (rows.length === 0) {
      list.innerHTML = '<div class="empty">履歴がまだありません</div>';
      return;
    }
    list.innerHTML = rows.map(r => `
      <div class="history-row ${r.type === '入庫' ? 'in' : 'out'}">
        <div class="h-type ${r.type === '入庫' ? 'in' : 'out'}">${escapeHtml(r.type)}</div>
        <div class="h-main">
          <div class="h-name">${escapeHtml(r.name)}</div>
          <div class="h-detail">${escapeHtml(r.datetime)} · ${escapeHtml(r.operator)} ${r.note ? '· ' + escapeHtml(r.note) : ''}</div>
        </div>
        <div class="h-qty">${r.type === '入庫' ? '+' : '−'}${formatNum(r.qty)}</div>
      </div>
    `).join('');
  } catch (err) {
    list.innerHTML = `<div class="empty">読み込み失敗: ${escapeHtml(err.message)}</div>`;
  }
}

// === 商品追加フォーム ===
function populateAddCategories() {
  const sel = document.getElementById('addCategory');
  sel.innerHTML = state.categories
    .map(c => `<option value="${escapeAttr(c.name)}">${c.icon || ''} ${escapeHtml(c.name)}</option>`)
    .join('');
}

function bindAddForm() {
  document.getElementById('addForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const payload = { action: 'addItem' };
    fd.forEach((v, k) => payload[k] = v);
    try {
      await apiPost(payload);
      toast('商品を登録しました');
      e.target.reset();
      await loadStock();
      // 在庫タブに戻る
      document.querySelector('.tab[data-tab="stock"]').click();
    } catch (err) {
      toast('失敗: ' + err.message, true);
    }
  });
}

// === ユーティリティ ===
function toast(msg, isError = false) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.remove('hidden', 'error');
  if (isError) t.classList.add('error');
  setTimeout(() => t.classList.add('hidden'), 2600);
}

function formatNum(n) {
  const num = Number(n);
  if (Number.isInteger(num)) return num.toString();
  return num.toFixed(2).replace(/\.?0+$/, '');
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }
