/* === Curry Stock - フロント JS === */

// ▼▼▼ ここを自分の Apps Script Web App URL に書き換える ▼▼▼
const API_URL = 'https://script.google.com/macros/s/AKfycbwakrB-QucMZKmY1u736DtS2OISpyVnllPDzHsL_mNwItRzYDqkBRy3MIAIAOANrpNy0A/exec';
// ▲▲▲ デプロイ後に取得する URL を貼り付ける ▲▲▲

// === 状態 ===
const state = {
  items: [],
  categories: [],
  locations: [],
  config: {},
  activeCategory: 'all',
  activeLocation: 'all', // 'all' | 拠点名
  searchQuery: '',
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
  bindSearch();
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
      if (target === 'add') {
        populateAddCategories();
        renderAddStockInputs();
      }
    });
  });
}

// === 検索 ===
function bindSearch() {
  const box = document.getElementById('searchBox');
  const clear = document.getElementById('searchClear');
  box.addEventListener('input', () => {
    state.searchQuery = box.value;
    clear.classList.toggle('hidden', !box.value);
    renderItems();
  });
  clear.addEventListener('click', () => {
    box.value = '';
    state.searchQuery = '';
    clear.classList.add('hidden');
    renderItems();
    box.focus();
  });
}

// === 在庫一覧読み込み ===
async function loadStock() {
  try {
    const data = await apiGet('list');
    state.items = data.items || [];
    state.categories = data.categories || [];
    state.locations = data.locations || [];
    state.config = data.config || {};
    renderLocationTabs();
    renderCategoryTabs();
    renderItems();
    renderAlertBar();
  } catch (err) {
    document.getElementById('itemsList').innerHTML =
      `<div class="empty">読み込み失敗: ${escapeHtml(err.message)}<br><br>API_URL を app.js で正しく設定したか確認してください。</div>`;
  }
}

// === 拠点タブ ===
function renderLocationTabs() {
  const wrap = document.getElementById('locationTabs');
  const items = [{ name: 'all', label: '全拠点合計', icon: '🏯' }]
    .concat(state.locations.map(l => ({
      name: l.name,
      label: l.name,
      icon: l.type === '倉庫' ? '🏭' : '🏠'
    })));
  wrap.innerHTML = items.map(l =>
    `<button class="loc-pill ${state.activeLocation === l.name ? 'active' : ''}" data-loc="${escapeAttr(l.name)}">
       <span class="loc-ico">${l.icon}</span>${escapeHtml(l.label)}
     </button>`
  ).join('');
  wrap.querySelectorAll('.loc-pill').forEach(p => {
    p.addEventListener('click', () => {
      state.activeLocation = p.dataset.loc;
      renderLocationTabs();
      renderItems();
    });
  });
}

// === カテゴリータブ ===
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

// === 商品カード描画 ===
function getStockForActiveLoc(item) {
  if (state.activeLocation === 'all') return item.total;
  return Number(item.stocks[state.activeLocation]) || 0;
}

// 在庫ステータス判定: 余裕(ok) / 注意(warn) / 不足(danger)
function getStockStatus(item) {
  if (item.minStock <= 0) return null; // 最低在庫未設定
  if (item.total <= item.minStock) return 'danger';
  if (item.total <= item.minStock * 1.5) return 'warn';
  return 'ok';
}

function statusBadgeHtml(status) {
  if (!status) return '';
  const labels = { ok: '余裕', warn: '注意', danger: '不足' };
  return `<span class="status-badge ${status}">${labels[status]}</span>`;
}

function renderItems() {
  const list = document.getElementById('itemsList');
  let items = state.items;
  if (state.activeCategory !== 'all') {
    items = items.filter(i => i.category === state.activeCategory);
  }
  const q = state.searchQuery.trim().toLowerCase();
  if (q) {
    items = items.filter(i =>
      String(i.name || '').toLowerCase().includes(q) ||
      String(i.code || '').toLowerCase().includes(q)
    );
  }
  if (items.length === 0) {
    list.innerHTML = q
      ? `<div class="empty">「${escapeHtml(q)}」に一致する商品がありません</div>`
      : '<div class="empty">商品がありません</div>';
    return;
  }
  const isAllLoc = state.activeLocation === 'all';
  list.innerHTML = items.map(it => {
    const status = getStockStatus(it);
    const showStock = getStockForActiveLoc(it);
    const totalLabel = isAllLoc ? '合計' : escapeHtml(state.activeLocation);

    const stocksGrid = isAllLoc ? `
      <div class="stocks-grid">
        ${state.locations.map(l => {
          const v = Number(it.stocks[l.name]) || 0;
          const isWarehouse = l.type === '倉庫';
          const isZero = v === 0;
          return `<div class="stock-cell ${isWarehouse ? 'warehouse' : ''} ${isZero ? 'zero' : ''}">
            <div class="stock-cell-name">${escapeHtml(l.name)}</div>
            <div class="stock-cell-num">${formatNum(v)}</div>
          </div>`;
        }).join('')}
      </div>
    ` : '';

    const cardClass = [
      'item-card',
      !isAllLoc ? 'single-loc' : '',
      status === 'danger' ? 'danger' : '',
      status === 'warn' ? 'warn' : ''
    ].filter(Boolean).join(' ');

    return `
      <div class="${cardClass}" data-code="${escapeAttr(it.code)}">
        <div class="item-header">
          <div class="item-info">
            <div class="item-name">${escapeHtml(it.name)}</div>
            <div class="item-meta">${escapeHtml(it.code)} · ${escapeHtml(it.category)}${it.minStock > 0 ? ' · 最低 ' + formatNum(it.minStock) + escapeHtml(it.unit) : ''}</div>
          </div>
          <div class="item-total">
            ${statusBadgeHtml(status)}
            <div>
              <span class="item-total-num">${formatNum(showStock)}</span>
              <span class="item-total-unit">${escapeHtml(it.unit)}</span>
            </div>
            <div class="item-total-label">${totalLabel}</div>
          </div>
        </div>
        ${stocksGrid}
      </div>`;
  }).join('');
  list.querySelectorAll('.item-card').forEach(card => {
    card.addEventListener('click', () => openModal(card.dataset.code));
  });
}

function renderAlertBar() {
  const bar = document.getElementById('alertBar');
  const lows = state.items.filter(i => i.minStock > 0 && i.total <= i.minStock);
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

  // 拠点別在庫の表示(横並び 4拠点 + 合計の 5セル grid)
  const stocksDiv = document.getElementById('modalStocks');
  const cells = state.locations.map(l => {
    const v = Number(item.stocks[l.name]) || 0;
    return `<div class="modal-stock-row"><span>${escapeHtml(l.name)}</span><strong>${formatNum(v)}<em>${escapeHtml(item.unit)}</em></strong></div>`;
  }).join('');
  stocksDiv.innerHTML = cells + `<div class="modal-stock-row total"><span>合計</span><strong>${formatNum(item.total)}<em>${escapeHtml(item.unit)}</em></strong></div>`;

  // 拠点セレクト(現在のフィルタ拠点を初期選択。allなら最初の拠点)
  const locSel = document.getElementById('modalLocation');
  locSel.innerHTML = state.locations.map(l =>
    `<option value="${escapeAttr(l.name)}">${escapeHtml(l.name)}</option>`
  ).join('');
  if (state.activeLocation !== 'all' && state.locations.find(l => l.name === state.activeLocation)) {
    locSel.value = state.activeLocation;
  }

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
  const location = document.getElementById('modalLocation').value;
  if (!location) {
    toast('拠点を選んでください', true);
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
      code, location, type, qty, operator, note
    });
    toast(`${type} OK: ${result.name} (${result.location}) → ${formatNum(result.stockAfter)}`);
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
          <div class="h-name">
            ${escapeHtml(r.name)}
            ${r.location ? `<span class="h-loc">${escapeHtml(r.location)}</span>` : ''}
          </div>
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

function renderAddStockInputs() {
  const div = document.getElementById('addStockInputs');
  div.innerHTML = state.locations.map(l => `
    <label class="stock-input-cell">
      <span>${l.type === '倉庫' ? '🏭' : '🏠'} ${escapeHtml(l.name)}</span>
      <input name="stock_${escapeAttr(l.name)}" type="number" min="0" step="0.01" value="0" />
    </label>
  `).join('');
}

function bindAddForm() {
  document.getElementById('addForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const stocks = {};
    state.locations.forEach(l => {
      const key = 'stock_' + l.name;
      stocks[l.name] = Number(fd.get(key)) || 0;
      fd.delete(key);
    });
    const payload = { action: 'addItem', stocks };
    fd.forEach((v, k) => payload[k] = v);
    try {
      await apiPost(payload);
      toast('商品を登録しました');
      e.target.reset();
      await loadStock();
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
  if (!isFinite(num)) return '0';
  if (Number.isInteger(num)) return num.toString();
  return num.toFixed(2).replace(/\.?0+$/, '');
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }
