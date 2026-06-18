// ── SpendSense App Logic ──

const CATEGORIES = [
  { id: 'food',      label: 'Food',       icon: '🍜' },
  { id: 'transport', label: 'Transport',  icon: '🚌' },
  { id: 'shopping',  label: 'Shopping',   icon: '🛍️' },
  { id: 'health',    label: 'Health',     icon: '💊' },
  { id: 'bills',     label: 'Bills',      icon: '⚡' },
  { id: 'rent',      label: 'Rent',       icon: '🏠' },
  { id: 'entertainment', label: 'Fun',    icon: '🎮' },
  { id: 'education', label: 'Education',  icon: '📚' },
  { id: 'other',     label: 'Other',      icon: '📦' },
];

let expenses = [];
let budget   = 0;
let selectedCat = '';

// ── Backend Configuration and Handlers ──
const API_BASE = (
  window.location.hostname === 'localhost' && window.location.port === '8080'
) ? 'http://localhost:5000' : (
  (window.location.hostname === 'localhost' || window.location.hostname === '') ? 'http://10.0.2.2:5000' : 'http://localhost:5000'
);

let isBackendOnline = false;

async function checkBackend() {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 1200);
    const res = await fetch(`${API_BASE}/`, { method: 'GET', signal: controller.signal });
    clearTimeout(timeoutId);
    isBackendOnline = res.ok;
  } catch (e) {
    isBackendOnline = false;
  }
}

async function fetchBudget() {
  await checkBackend();
  if (isBackendOnline) {
    try {
      const res = await fetch(`${API_BASE}/api/budget`);
      const data = await res.json();
      budget = parseFloat(data.budget) || 0;
      localStorage.setItem('spendsense_budget', budget);
    } catch (e) {
      console.warn('Failed to fetch budget from backend, using local:', e);
      budget = parseFloat(localStorage.getItem('spendsense_budget') || '0');
    }
  } else {
    budget = parseFloat(localStorage.getItem('spendsense_budget') || '0');
  }
}

async function fetchExpenses() {
  await checkBackend();
  if (isBackendOnline) {
    try {
      const res = await fetch(`${API_BASE}/api/expenses`);
      expenses = await res.json();
      localStorage.setItem('spendsense_expenses', JSON.stringify(expenses));
    } catch (e) {
      console.warn('Failed to fetch expenses from backend, using local:', e);
      expenses = JSON.parse(localStorage.getItem('spendsense_expenses') || '[]');
    }
  } else {
    expenses = JSON.parse(localStorage.getItem('spendsense_expenses') || '[]');
  }
}

// ── Init ──
async function init() {
  document.getElementById('todayDate').textContent =
    new Date().toLocaleDateString('en-IN', { weekday:'short', day:'numeric', month:'short', year:'numeric' });
  document.getElementById('currentMonth').textContent =
    new Date().toLocaleDateString('en-IN', { month:'long', year:'numeric' });

  const today = new Date().toISOString().split('T')[0];
  document.getElementById('expDate').value = today;

  buildCategoryPicker();
  buildFilterDropdown();

  // Load from database if available
  await fetchBudget();
  await fetchExpenses();

  renderDashboard();
  renderHistory();
  renderBudgetPage();
}

// ── Category Picker ──
function buildCategoryPicker() {
  const picker = document.getElementById('categoryPicker');
  picker.innerHTML = '';
  CATEGORIES.forEach(cat => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'cat-option' + (selectedCat === cat.id ? ' selected' : '');
    btn.textContent = cat.icon + ' ' + cat.label;
    btn.onclick = () => selectCat(cat.id);
    picker.appendChild(btn);
  });
}

function selectCat(id) {
  selectedCat = id;
  buildCategoryPicker();
}

// ── Filter Dropdown ──
function buildFilterDropdown() {
  const sel = document.getElementById('filterCat');
  sel.innerHTML = '<option value="">All Categories</option>';
  CATEGORIES.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c.id; opt.textContent = c.icon + ' ' + c.label;
    sel.appendChild(opt);
  });
}

// ── Add Expense ──
async function addExpense() {
  const desc   = document.getElementById('expDesc').value.trim();
  const amount = parseFloat(document.getElementById('expAmount').value);
  const date   = document.getElementById('expDate').value;
  const note   = document.getElementById('expNote').value.trim();

  const fb = document.getElementById('formFeedback');

  if (!desc) { showFeedback(fb, 'Please enter a description.', 'error'); return; }
  if (!amount || amount <= 0) { showFeedback(fb, 'Please enter a valid amount.', 'error'); return; }
  if (!date) { showFeedback(fb, 'Please select a date.', 'error'); return; }
  if (!selectedCat) { showFeedback(fb, 'Please select a category.', 'error'); return; }

  const expense = {
    id: Date.now(),
    desc, amount, date, note,
    category: selectedCat,
  };

  expenses.unshift(expense);
  save();

  await checkBackend();
  if (isBackendOnline) {
    try {
      await fetch(`${API_BASE}/api/expenses`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(expense)
      });
      // fetch latest list to be sure
      const res = await fetch(`${API_BASE}/api/expenses`);
      expenses = await res.json();
      save();
    } catch (e) {
      console.error('Failed to save to backend, saved locally:', e);
    }
  }

  // reset
  document.getElementById('expDesc').value   = '';
  document.getElementById('expAmount').value = '';
  document.getElementById('expNote').value   = '';
  selectedCat = '';
  buildCategoryPicker();

  showFeedback(fb, '✓ Expense recorded!', 'success');
  renderDashboard();
  renderHistory();
  renderBudgetPage();
}

function showFeedback(el, msg, type) {
  el.textContent = msg;
  el.className = 'form-feedback ' + type;
  setTimeout(() => { el.textContent = ''; el.className = 'form-feedback'; }, 2500);
}

// ── Delete ──
async function deleteExpense(id) {
  expenses = expenses.filter(e => e.id !== id);
  save();

  await checkBackend();
  if (isBackendOnline) {
    try {
      await fetch(`${API_BASE}/api/expenses/${id}`, { method: 'DELETE' });
    } catch (e) {
      console.error('Failed to delete from backend:', e);
    }
  }

  renderDashboard();
  renderHistory();
  renderBudgetPage();
}

async function clearAll() {
  if (!confirm('Delete all expenses? This cannot be undone.')) return;
  expenses = [];
  save();

  await checkBackend();
  if (isBackendOnline) {
    try {
      await fetch(`${API_BASE}/api/expenses`, { method: 'DELETE' });
    } catch (e) {
      console.error('Failed to clear from backend:', e);
    }
  }

  renderDashboard();
  renderHistory();
  renderBudgetPage();
}

// ── Budget ──
async function setBudget() {
  const val = parseFloat(document.getElementById('budgetInput').value);
  const fb  = document.getElementById('budgetFeedback');
  if (!val || val <= 0) { showFeedback(fb, 'Please enter a valid budget.', 'error'); return; }
  budget = val;
  localStorage.setItem('spendsense_budget', val);

  await checkBackend();
  if (isBackendOnline) {
    try {
      await fetch(`${API_BASE}/api/budget`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ budget: val })
      });
    } catch (e) {
      console.error('Failed to set budget in backend:', e);
    }
  }

  showFeedback(fb, '✓ Budget set to ₹' + fmt(val), 'success');
  document.getElementById('budgetInput').value = '';
  renderDashboard();
  renderBudgetPage();
}

// ── Helpers ──
function save() { localStorage.setItem('spendsense_expenses', JSON.stringify(expenses)); }

function fmt(n) {
  return Number(n).toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function getCat(id) { return CATEGORIES.find(c => c.id === id) || { label: id, icon: '📦' }; }

function thisMonthExpenses() {
  const now = new Date();
  return expenses.filter(e => {
    const d = new Date(e.date);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  });
}

function catTotals(list) {
  const totals = {};
  list.forEach(e => {
    totals[e.category] = (totals[e.category] || 0) + e.amount;
  });
  return totals;
}

// ── Dashboard ──
function renderDashboard() {
  const monthly = thisMonthExpenses();
  const total   = monthly.reduce((s,e) => s+e.amount, 0);
  const totals  = catTotals(monthly);

  document.getElementById('totalSpent').textContent  = '₹' + fmt(total);
  document.getElementById('monthCount').textContent  = monthly.length;

  // highest category
  let highCat = '—';
  if (Object.keys(totals).length > 0) {
    const topId = Object.entries(totals).sort((a,b)=>b[1]-a[1])[0][0];
    highCat = getCat(topId).icon + ' ' + getCat(topId).label;
  }
  document.getElementById('highestCat').textContent = highCat;

  // budget
  if (budget > 0) {
    const rem = budget - total;
    document.getElementById('remainingBudget').textContent = '₹' + fmt(Math.max(0, rem));
    document.getElementById('budgetStatus').textContent    = rem < 0 ? '⚠ over budget!' : 'remaining';
    document.getElementById('spentVsBudget').textContent   = 'of ₹' + fmt(budget) + ' budget';

    const pct = Math.min((total/budget)*100, 100);
    const fill = document.getElementById('budgetBarFill');
    fill.style.width = pct + '%';
    fill.classList.toggle('danger', pct >= 85);
    document.getElementById('budgetBarSpent').textContent = '₹' + fmt(total) + ' spent';
    document.getElementById('budgetBarTotal').textContent = 'of ₹' + fmt(budget);
    document.getElementById('budgetBarSection').style.display = 'block';
  } else {
    document.getElementById('remainingBudget').textContent = '—';
    document.getElementById('budgetStatus').textContent    = 'set a budget';
    document.getElementById('spentVsBudget').textContent   = 'no budget set';
    document.getElementById('budgetBarSection').style.display = 'none';
  }

  // category grid
  const grid = document.getElementById('categoryGrid');
  if (Object.keys(totals).length === 0) {
    grid.innerHTML = '<div class="empty-state">No expenses yet. Add your first one!</div>';
  } else {
    grid.innerHTML = '';
    Object.entries(totals).sort((a,b)=>b[1]-a[1]).forEach(([id, amt]) => {
      const cat = getCat(id);
      const count = monthly.filter(e=>e.category===id).length;
      const card = document.createElement('div');
      card.className = 'cat-card';
      card.innerHTML = `
        <div class="cat-icon">${cat.icon}</div>
        <div class="cat-name">${cat.label}</div>
        <div class="cat-amount">₹${fmt(amt)}</div>
        <div class="cat-count">${count} transaction${count!==1?'s':''}</div>
      `;
      grid.appendChild(card);
    });
  }

  // recent (last 5)
  renderTxList(document.getElementById('recentList'), monthly.slice(0,5));
}

// ── History ──
function renderHistory() {
  const filter = document.getElementById('filterCat').value;
  let list = [...expenses];
  if (filter) list = list.filter(e => e.category === filter);
  renderTxList(document.getElementById('historyList'), list, true);
}

// ── Budget Page ──
function renderBudgetPage() {
  if (budget > 0) {
    document.getElementById('budgetInput').placeholder = 'Current: ₹' + fmt(budget);
  }
  const monthly = thisMonthExpenses();
  const totals  = catTotals(monthly);
  const maxAmt  = Math.max(...Object.values(totals), 1);

  const detail = document.getElementById('categoryDetail');
  if (Object.keys(totals).length === 0) {
    detail.innerHTML = '<div class="empty-state">No data yet.</div>';
    return;
  }
  detail.innerHTML = '';
  Object.entries(totals).sort((a,b)=>b[1]-a[1]).forEach(([id, amt]) => {
    const cat   = getCat(id);
    const count = monthly.filter(e=>e.category===id).length;
    const pct   = (amt/maxAmt)*100;
    const row   = document.createElement('div');
    row.className = 'cat-detail-row';
    row.innerHTML = `
      <div class="cat-detail-icon">${cat.icon}</div>
      <div class="cat-detail-info">
        <div class="cat-detail-name">${cat.label}</div>
        <div class="cat-detail-bar-track">
          <div class="cat-detail-bar-fill" style="width:${pct}%"></div>
        </div>
      </div>
      <div>
        <div class="cat-detail-amount">₹${fmt(amt)}</div>
        <div class="cat-detail-count">${count} tx</div>
      </div>
    `;
    detail.appendChild(row);
  });
}

// ── Render Transaction List ──
function renderTxList(container, list, showDelete = false) {
  if (list.length === 0) {
    container.innerHTML = '<div class="empty-state">No transactions found.</div>';
    return;
  }
  container.innerHTML = '';
  list.forEach(e => {
    const cat  = getCat(e.category);
    const item = document.createElement('div');
    item.className = 'tx-item';
    item.innerHTML = `
      <div class="tx-icon">${cat.icon}</div>
      <div class="tx-info">
        <div class="tx-desc">${e.desc}</div>
        <div class="tx-meta">${cat.label} · ${new Date(e.date).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}${e.note ? ' · ' + e.note : ''}</div>
      </div>
      <div class="tx-amount">−₹${fmt(e.amount)}</div>
      <button class="tx-delete" onclick="deleteExpense(${e.id})" title="Delete">✕</button>
    `;
    container.appendChild(item);
  });
}

// ── Navigation ──
function showView(name) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('view-' + name).classList.add('active');
  document.querySelector(`[data-view="${name}"]`).classList.add('active');

  if (name === 'history') renderHistory();
  if (name === 'budget')  renderBudgetPage();

  // Close sidebar on mobile devices
  toggleSidebar(false);
}

function toggleSidebar(isOpen) {
  const sidebar = document.querySelector('.sidebar');
  const overlay = document.getElementById('sidebarOverlay');
  if (sidebar && overlay) {
    if (isOpen) {
      sidebar.classList.add('open');
      overlay.classList.add('open');
    } else {
      sidebar.classList.remove('open');
      overlay.classList.remove('open');
    }
  }
}

// ── Start ──
init();
