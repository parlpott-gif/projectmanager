// ─── 状态 ────────────────────────────────────────────
let allProjects = [];
let currentFilter = 'all';
let currentProject = null;
let expandedProjectName = null; // 当前展开的项目名

let allParts = [];
let partsFilter = 'all';
let currentPage = 'projects';

// ─── 初始化 ──────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadProjects();
  document.getElementById('edit-date').value = new Date().toISOString().split('T')[0];
});

// ─── 页面切换 ─────────────────────────────────────────
function switchPage(page) {
  const previousPage = currentPage;  // 保存上一个页面，用于 export-history 跳转判断
  currentPage = page;
  document.getElementById('page-projects').style.display = page === 'projects' ? '' : 'none';
  document.getElementById('page-parts').style.display = page === 'parts' ? '' : 'none';
  document.getElementById('page-stats').style.display = page === 'stats' ? '' : 'none';
  document.getElementById('page-quick-export').style.display = page === 'quick-export' ? '' : 'none';
  document.getElementById('page-export-history').style.display = page === 'export-history' ? '' : 'none';

  // 工具栏控制 - 同一位置轮换显示三个搜索区块
  const filterContainer = document.getElementById('filter-container');
  if (filterContainer) {
    filterContainer.style.display = ['projects', 'parts', 'quick-export'].includes(page) ? '' : 'none';
  }
  const projFilterBlock = document.getElementById('projects-filter-block');
  const partsFilterBlock = document.getElementById('parts-filter-block');
  const exportFilterBlock = document.getElementById('export-filter-block');
  if (projFilterBlock) projFilterBlock.style.display = page === 'projects' ? '' : 'none';
  if (partsFilterBlock) partsFilterBlock.style.display = page === 'parts' ? '' : 'none';
  if (exportFilterBlock) exportFilterBlock.style.display = page === 'quick-export' ? '' : 'none';

  // Tab 激活状态
  document.getElementById('tab-projects').classList.toggle('active', page === 'projects');
  document.getElementById('tab-parts').classList.toggle('active', page === 'parts');
  document.getElementById('tab-stats').classList.toggle('active', page === 'stats');
  document.getElementById('tab-quick-export').classList.toggle('active', page === 'quick-export');
  document.getElementById('tab-export-history').classList.toggle('active', page === 'export-history');

  // 页面初始化
  if (page === 'parts') loadParts();
  if (page === 'stats') loadDailyStats();
  if (page === 'quick-export') initUniversalExport();
  if (page === 'export-history') {
    // 如果是从导出页切换来的，传入当前导出类型；否则不筛选
    const exportType = previousPage === 'quick-export' ? universalExportState.selectedExportType : null;
    loadExportHistory(exportType);
  }
}

// ─── 数据加载 ─────────────────────────────────────────
async function loadProjects() {
  showLoading(true);
  try {
    const res = await fetch('/api/projects');
    if (!res.ok) throw new Error('请求失败');
    allProjects = await res.json();
    renderTable();
    loadStats();
  } catch (e) {
    showToast('读取失败：' + e.message, 'danger');
  } finally {
    showLoading(false);
  }
}

async function updateAllCosts() {
  const btn = document.getElementById('btn-update-costs');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>计算中...'; }
  try {
    const res = await fetch('/api/costs/update', { method: 'POST' });
    const data = await res.json();
    if (!data.ok) { showToast('更新失败：' + (data.error || '未知错误'), 'danger'); return; }
    showToast(`硬件成本已更新：${data.updated}/${data.total} 个项目`, 'success');
    await loadProjects();
  } catch (e) {
    showToast('更新失败：' + e.message, 'danger');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-calculator me-1"></i>更新成本'; }
  }
}

async function loadStats() {
  try {
    const res = await fetch('/api/stats');
    const s = await res.json();
    document.getElementById('stat-total').textContent = s.total;
    document.getElementById('stat-price').textContent = formatMoney(s.total_price);
    document.getElementById('stat-prepay').textContent = formatMoney(s.total_prepay);
  } catch (e) { /* 忽略 */ }
}

async function syncData() {
  showToast('正在同步总表...', 'info');
  await loadProjects();
  showToast('同步完成', 'success');
}

// ─── 阶段排序权重 ─────────────────────────────────────
const STATUS_ORDER = { '待绘制': 0, '待打板': 1, '等板': 2, '待交付': 3, '已完成': 4 };

// ─── 渲染表格 ──────────────────────────────────────────
function renderTable() {
  const wrap = document.getElementById('project-table-wrap');
  const empty = document.getElementById('empty-state');
  const tbody = document.getElementById('project-tbody');
  const keyword = (document.getElementById('search-input') || {value:''}).value.toLowerCase();

  let list = allProjects;

  if (currentFilter !== 'all') {
    list = list.filter(p => (p.status || '待绘制') === currentFilter);
  }
  if (keyword) {
    list = list.filter(p => {
      // 支持按项目名称、分类、交付类型搜索
      const nameMatch = p.name.toLowerCase().includes(keyword);
      const categoryMatch = (p.category || '').toLowerCase().includes(keyword);
      const deliverableMatch = (p.deliverable || '').toLowerCase().includes(keyword);

      // 支持按编号搜索（可以输入"001"或"1"）
      let idMatch = false;
      if (p.id) {
        const idStr = String(p.id).padStart(3, '0');  // "1" -> "001"
        const idNumStr = String(p.id);  // "1"
        idMatch = idStr.includes(keyword) || idNumStr.includes(keyword);
      }

      return nameMatch || categoryMatch || deliverableMatch || idMatch;
    });
  }

  // 按阶段权重排序：待绘制 → 待打板 → 等板 → 待交付 → 已完成
  list = list.slice().sort((a, b) => {
    const wa = STATUS_ORDER[a.status || '待绘制'] ?? 99;
    const wb = STATUS_ORDER[b.status || '待绘制'] ?? 99;
    return wa - wb;
  });

  if (list.length === 0) {
    wrap.classList.add('d-none');
    empty.classList.remove('d-none');
    return;
  }

  wrap.classList.remove('d-none');
  empty.classList.add('d-none');

  tbody.innerHTML = list.map(p => {
    let html = renderRow(p);
    if (expandedProjectName === p.name) {
      html += renderExpandedRow(p);
    }
    return html;
  }).join('');
}

const STATUS_LIST = ['待绘制', '待打板', '等板', '待交付', '已完成'];

function renderRow(p) {
  const status = p.status || '待绘制';
  const price = p.price ? `<span class="td-price">¥${p.price}</span>` : '<span class="text-muted">—</span>';
  const prepay = p.prepay ? `<span class="td-prepay">¥${p.prepay}</span>` : '<span class="text-muted">—</span>';
  const date = p.date ? `<span class="td-date">${p.date}</span>` : '<span class="text-muted">—</span>';
  const deliverable = p.deliverable ? escHtml(p.deliverable) : '<span class="text-muted">—</span>';
  const costHtml = p.hw_cost != null
    ? `<span class="td-cost">¥${p.hw_cost}</span>`
    : `<span class="text-muted">—</span>`;

  const statusOptions = STATUS_LIST.map(s =>
    `<option value="${s}"${s === status ? ' selected' : ''}>${s}</option>`
  ).join('');

  const isExpanded = expandedProjectName === p.name;

  // 编号显示
  const idDisplay = p.id ? `[${String(p.id).padStart(3, '0')}]` : '—';

  return `
<tr data-name="${escAttr(p.name)}" class="project-row ${isExpanded ? 'expanded' : ''}" style="cursor:pointer" onclick="toggleProjectExpand('${escAttr(p.name)}')">
  <td style="text-align:center; font-weight:600; color:var(--accent)">${idDisplay}</td>
  <td class="td-name" title="${escAttr(p.name)}">${escHtml(p.name)}</td>
  <td>${deliverable}</td>
  <td>${price}</td>
  <td>${prepay}</td>
  <td>${costHtml}</td>
  <td>${date}</td>
  <td onclick="event.stopPropagation()">
    <select class="inline-status-select status-${status}" data-name="${escAttr(p.name)}"
      onchange="inlineChangeStatus(this, '${escAttr(p.name)}')">
      ${statusOptions}
    </select>
  </td>
  <td onclick="event.stopPropagation()">
    <button class="btn-detail" onclick="openSidebar('${escAttr(p.name)}')">
      <i class="bi bi-pencil-square"></i>
    </button>
    <button class="btn-detail ms-1" title="打开文件夹" onclick="openFolderByName('${escAttr(p.name)}')">
      <i class="bi bi-folder2-open"></i>
    </button>
  </td>
</tr>`;
}

// ─── 展开行渲染 ────────────────────────────────────────
function renderExpandedRow(p) {
  const comps = parseLines(p.components);
  const funcs = parseLines(p.functions);

  const compsHtml = comps.length
    ? comps.map(c => `<div class="comp-row">${escHtml(c)}</div>`).join('')
    : '<div class="text-muted small">暂无数据</div>';

  const funcsHtml = funcs.length
    ? funcs.map(f => `<div class="func-row">${escHtml(f)}</div>`).join('')
    : '<div class="text-muted small">暂无数据</div>';

  return `
<tr class="expand-row">
  <td colspan="8">
    <div class="expand-content">
      <div class="expand-section">
        <h6 class="expand-title"><i class="bi bi-cpu me-2"></i>元器件清单</h6>
        <div class="component-list">${compsHtml}</div>
      </div>
      <div class="expand-section">
        <h6 class="expand-title"><i class="bi bi-list-check me-2"></i>功能列表</h6>
        <div class="function-list">${funcsHtml}</div>
      </div>
    </div>
  </td>
</tr>`;
}

// ─── 切换项目展开 ──────────────────────────────────────
function toggleProjectExpand(name) {
  if (expandedProjectName === name) {
    expandedProjectName = null;
  } else {
    expandedProjectName = name;
  }
  renderTable();
}

// ─── 筛选 ─────────────────────────────────────────────
function setFilter(btn, status) {
  document.querySelectorAll('#projects-filter-block .filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  currentFilter = status;
  renderTable();
}

function filterProjects() { renderTable(); }

// ─── 行内修改阶段 ──────────────────────────────────────
async function inlineChangeStatus(select, name) {
  const newStatus = select.value;
  const p = allProjects.find(x => x.name === name);
  if (!p) return;
  const oldStatus = p.status || '待绘制';
  if (newStatus === oldStatus) return;

  // 立即更新样式
  select.className = `inline-status-select status-${newStatus}`;

  try {
    const body = {
      deliverable: p.deliverable || '',
      price: p.price || '',
      prepay: p.prepay || '',
      date: p.date || '',
      status: newStatus,
      note: p.note || '',
    };
    const res = await fetch('/api/projects/' + encodeURIComponent(p.name), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error('保存失败');
    p.status = newStatus;
    showToast(`已更新为「${newStatus}」`, 'success');
    // 重新排序渲染
    renderTable();
  } catch (e) {
    showToast('保存失败：' + e.message, 'danger');
    // 回滚
    select.value = oldStatus;
    select.className = `inline-status-select status-${oldStatus}`;
  }
}

// ─── 侧边栏 ───────────────────────────────────────────
function openSidebar(name) {
  const p = allProjects.find(x => x.name === name);
  if (!p) return;
  currentProject = p;

  document.getElementById('sidebar-name').textContent = p.name;
  const status = p.status || '待绘制';
  const badge = document.getElementById('sidebar-status-badge');
  badge.textContent = status;
  badge.className = `badge status-badge status-${status}`;

  document.getElementById('edit-deliverable').value = p.deliverable || '';
  document.getElementById('edit-price').value = p.price || '';
  document.getElementById('edit-prepay').value = p.prepay || '';
  document.getElementById('edit-date').value = p.date || new Date().toISOString().split('T')[0];
  document.getElementById('edit-status').value = status;

  const comps = parseLines(p.components);
  document.getElementById('sidebar-components').innerHTML = comps.length
    ? comps.map(c => `<div class="comp-row">${escHtml(c)}</div>`).join('')
    : '<div class="text-muted small">暂无数据</div>';

  const funcs = parseLines(p.functions);
  document.getElementById('sidebar-functions').innerHTML = funcs.length
    ? funcs.map(f => `<div class="func-row">${escHtml(f)}</div>`).join('')
    : '<div class="text-muted small">暂无数据</div>';

  // 根据项目阶段自动选择上传目录
  const autoSelectUploadDir = () => {
    const status = p.status || '待绘制';
    let selectedDir = '硬件';  // 默认值

    // 根据阶段自动选择
    if (status === '等板' || status === '待绘制' || status === '待打板') {
      selectedDir = '硬件';  // 硬件设计阶段
    } else if (status === '待交付' || status === '已完成') {
      selectedDir = '演示';  // 交付/演示阶段
    }

    // 设置单选框
    document.getElementById('upload-dir-hardware').checked = selectedDir === '硬件';
    document.getElementById('upload-dir-demo').checked = selectedDir === '演示';
  };
  autoSelectUploadDir();

  // 渲染文件列表（支持子文件夹路径显示）
  renderSidebarFiles(p);

  document.getElementById('sidebar-overlay').classList.add('show');
  document.getElementById('sidebar').classList.add('show');
}

function closeSidebar() {
  document.getElementById('sidebar-overlay').classList.remove('show');
  document.getElementById('sidebar').classList.remove('show');
  currentProject = null;
}

// ─── 侧边栏文件列表渲染 ─────────────────────────────────
function renderSidebarFiles(p) {
  const fileEl = document.getElementById('sidebar-files');
  if (!p.files || p.files.length === 0) {
    fileEl.innerHTML = `<div class="text-muted small">${p.folder ? '文件夹存在但无文件' : '项目文件夹不存在'}</div>`;
    selectedFilesForExport = [];
    updateExportButtonState();
    return;
  }

  // 将文件分组：有子文件夹的归组，根目录的归入"根目录"
  const groups = {};
  for (const f of p.files) {
    const slashIdx = f.indexOf('/');
    if (slashIdx !== -1) {
      const sub = f.substring(0, slashIdx);
      const fname = f.substring(slashIdx + 1);
      if (!groups[sub]) groups[sub] = [];
      groups[sub].push(fname);
    } else {
      if (!groups['']) groups[''] = [];
      groups[''].push(f);
    }
  }

  let html = '';
  selectedFilesForExport = [];  // 重置选择列表

  // 先渲染根目录文件（通常不需要导出）
  if (groups['']) {
    html += groups[''].map(f =>
      `<div class="file-row"><i class="bi ${getFileIcon(f)} me-1"></i><span class="file-name">${escHtml(f)}</span></div>`
    ).join('');
  }

  // 再渲染子文件夹（硬件素材、演示） - 这些可以导出
  for (const [sub, files] of Object.entries(groups)) {
    if (!sub) continue;
    html += `<div class="file-subfolder-label"><i class="bi bi-folder2 me-1"></i>${escHtml(sub)}</div>`;
    html += files.map(f => {
      const fullName = `${sub}/${f}`;  // 保存完整路径用于导出
      return `<div class="file-row file-row-sub">
        <input type="checkbox" class="file-checkbox" data-filename="${escHtml(fullName)}"
          onchange="updateExportSelection('${escHtml(fullName)}', this.checked)">
        <i class="bi ${getFileIcon(f)} me-1"></i>
        <span class="file-name">${escHtml(f)}</span>
      </div>`;
    }).join('');
  }
  fileEl.innerHTML = html;
  updateExportButtonState();
}

// ─── 保存 ─────────────────────────────────────────────
async function saveProject() {
  if (!currentProject) return;
  const body = {
    deliverable: document.getElementById('edit-deliverable').value,
    price: document.getElementById('edit-price').value,
    prepay: document.getElementById('edit-prepay').value,
    date: document.getElementById('edit-date').value,
    status: document.getElementById('edit-status').value,
  };
  try {
    const res = await fetch('/api/projects/' + encodeURIComponent(currentProject.name), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error('保存失败');
    Object.assign(currentProject, body);
    showToast('已保存', 'success');
    renderTable();
    loadStats();
    const badge = document.getElementById('sidebar-status-badge');
    badge.textContent = body.status;
    badge.className = `badge status-badge status-${body.status}`;
  } catch (e) {
    showToast('保存失败：' + e.message, 'danger');
  }
}

// ─── 打开文件夹 ────────────────────────────────────────
async function openFolder() {
  if (!currentProject) return;
  await openFolderByName(currentProject.name);
}

async function openFolderByName(name) {
  try {
    const res = await fetch('/api/open_folder/' + encodeURIComponent(name), { method: 'POST' });
    const data = await res.json();
    if (!data.ok) showToast(data.error || '文件夹不存在', 'warning');
  } catch (e) { showToast('操作失败', 'danger'); }
}

// ─── 文件上传 ──────────────────────────────────────────
async function uploadFiles(inputEl) {
  if (!currentProject) return;
  const files = inputEl.files;
  if (!files || files.length === 0) return;

  const btn = document.getElementById('btn-upload-file');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>上传中...'; }

  // 获取选中的上传目录
  const uploadDir = document.querySelector('input[name="upload-dir"]:checked')?.value || '硬件';

  const formData = new FormData();
  for (const f of files) {
    formData.append('files', f);
  }
  formData.append('upload_dir', uploadDir);

  try {
    const res = await fetch('/api/projects/' + encodeURIComponent(currentProject.name) + '/upload', {
      method: 'POST',
      body: formData,
    });
    const data = await res.json();
    if (!data.ok) {
      showToast('上传失败：' + (data.error || '未知错误'), 'danger');
      return;
    }

    // 构建成功/失败提示
    const savedCount = data.saved.length;
    const errCount = data.errors ? data.errors.length : 0;
    let msg = `已上传 ${savedCount} 个文件到「${data.sub_folder}」`;
    if (errCount > 0) msg += `，${errCount} 个失败`;
    showToast(msg, savedCount > 0 ? 'success' : 'warning');

    // 刷新 currentProject.files 并重新渲染文件列表
    // 追加新文件（不重新请求整个 projects API，避免感知延迟）
    if (data.saved.length > 0) {
      if (!currentProject.files) currentProject.files = [];
      currentProject.files.push(...data.saved);
      renderSidebarFiles(currentProject);
    }
  } catch (e) {
    showToast('上传失败：' + e.message, 'danger');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-upload me-1"></i>上传文件'; }
    // 清空 input，使同一文件下次仍可触发 change 事件
    inputEl.value = '';
  }
}

// ─── 侧边栏折叠 ────────────────────────────────────────
function toggleSection(header) {
  const body = header.nextElementSibling;
  body.classList.toggle('show', !body.classList.contains('show'));
  header.classList.toggle('open', !header.classList.contains('open'));
}

// ─── 工具函数 ─────────────────────────────────────────
function parseLines(str) {
  if (!str) return [];
  return str.split(/\n/).map(s => s.trim()).filter(Boolean);
}

function formatMoney(n) {
  if (!n) return '0';
  return n.toLocaleString('zh-CN', { maximumFractionDigits: 0 });
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escAttr(str) {
  return String(str).replace(/'/g, '%27').replace(/"/g, '%22').replace(/\s/g, '_');
}

function getFileIcon(filename) {
  // filename 可能是 "硬件素材/xxx.pdf"，只取最后的文件名部分来判断图标
  const base = filename.includes('/') ? filename.split('/').pop() : filename;
  if (/\.(docx?)/i.test(base)) return 'bi-file-earmark-word text-primary';
  if (/\.md$/i.test(base)) return 'bi-file-earmark-text';
  if (/\.(xlsx?)/i.test(base)) return 'bi-file-earmark-excel text-success';
  if (/\.pdf$/i.test(base)) return 'bi-file-earmark-pdf text-danger';
  if (/\.(png|jpg|jpeg|gif|bmp)/i.test(base)) return 'bi-file-earmark-image text-warning';
  if (/\.(mp4|mov|avi|mkv)/i.test(base)) return 'bi-camera-video text-info';
  if (/\.zip$/i.test(base)) return 'bi-file-earmark-zip';
  return 'bi-file-earmark';
}

// ─── 器件库 ───────────────────────────────────────────

async function loadParts() {
  document.getElementById('parts-loading').style.display = 'block';
  document.getElementById('parts-table-wrap').style.display = 'none';
  document.getElementById('parts-empty').style.display = 'none';
  try {
    const res = await fetch('/api/parts');
    allParts = await res.json();
    renderParts();
  } catch (e) {
    showToast('器件库加载失败：' + e.message, 'danger');
  } finally {
    document.getElementById('parts-loading').style.display = 'none';
  }
}

function renderParts() {
  const tbody = document.getElementById('parts-tbody');
  const keyword = (document.getElementById('parts-search').value || '').toLowerCase();

  let list = allParts;
  if (partsFilter === 'shortage') list = list.filter(p => p.diff < 0);
  if (partsFilter === 'ok') list = list.filter(p => p.diff >= 0);
  if (keyword) list = list.filter(p => p.name.toLowerCase().includes(keyword));

  if (list.length === 0) {
    document.getElementById('parts-table-wrap').style.display = 'none';
    document.getElementById('parts-empty').style.display = 'block';
    return;
  }

  document.getElementById('parts-table-wrap').style.display = '';
  document.getElementById('parts-empty').style.display = 'none';

  tbody.innerHTML = list.map(p => renderPartRow(p)).join('');
}

function renderPartRow(p) {
  // 差额样式
  let diffClass = 'diff-zero', diffText = '—';
  if (p.demand > 0) {
    if (p.diff > 0) { diffClass = 'diff-ok'; diffText = `+${p.diff}`; }
    else if (p.diff === 0) { diffClass = 'diff-warn'; diffText = '0'; }
    else { diffClass = 'diff-bad'; diffText = String(p.diff); }
  }

  // 来源项目标签
  const projTags = p.projects.length
    ? `<div class="proj-tags">${p.projects.map(x => `<span class="proj-tag">${escHtml(x)}</span>`).join('')}</div>`
    : '<span class="text-muted small">—</span>';

  // 最近采购记录（最新一条）
  const last = p.purchases.length ? p.purchases[p.purchases.length-1] : null;
  const lastBuy = last
    ? `<div class="purchase-log"><span>${last.date}</span><span>+${last.qty}个</span>${last.unit_price ? `<span>¥${last.unit_price}/个</span>` : ''}</div>`
    : '';

  return `
<tr>
  <td style="font-weight:600;font-size:0.87rem">${escHtml(p.name)}</td>
  <td>
    <div class="td-stock-wrap" style="justify-content:center">
      <button class="btn-stock" onclick="adjustStock('${escAttr(p.name)}', -1)">−</button>
      <span class="stock-num" id="stock-${escAttr(p.name)}">${p.stock}</span>
      <button class="btn-stock" onclick="adjustStock('${escAttr(p.name)}', 1)">+</button>
    </div>
  </td>
  <td class="td-demand">${p.demand || '—'}</td>
  <td class="${diffClass}">${diffText}</td>
  <td>${projTags}</td>
  <td>
    <div class="purchase-wrap">
      <input type="number" class="form-control form-control-sm purchase-input"
        id="buy-${escAttr(p.name)}" min="1" placeholder="数量" autocomplete="off">
      <input type="number" class="form-control form-control-sm purchase-input"
        id="price-${escAttr(p.name)}" min="0" step="0.01" placeholder="单价" autocomplete="off">
      <button class="btn-buy" onclick="doPurchase('${escAttr(p.name)}')">入库</button>
    </div>
    ${lastBuy}
  </td>
</tr>`;
}

function updatePartRowInPlace(part) {
  // 局部更新单行：库存数字、差额单元格、最近采购记录，不重绘整表
  const key = escAttr(part.name);
  const stockEl = document.getElementById('stock-' + key);
  if (!stockEl) return;
  const tr = stockEl.closest('tr');
  if (!tr) return;

  // 库存
  stockEl.textContent = part.stock;

  // 差额（第4列，index=3）
  const diffTd = tr.cells[3];
  if (diffTd) {
    let diffClass = 'diff-zero', diffText = '—';
    if (part.demand > 0) {
      if (part.diff > 0)      { diffClass = 'diff-ok';   diffText = `+${part.diff}`; }
      else if (part.diff === 0){ diffClass = 'diff-warn'; diffText = '0'; }
      else                     { diffClass = 'diff-bad';  diffText = String(part.diff); }
    }
    diffTd.className = diffClass;
    diffTd.textContent = diffText;
  }

  // 最近采购记录（采购列内的 .purchase-log）
  const purchaseTd = tr.cells[5];
  if (purchaseTd && part.purchases.length) {
    const last = part.purchases[part.purchases.length - 1];
    const logHtml = `<div class="purchase-log"><span>${last.date}</span><span>+${last.qty}个</span>${last.unit_price ? `<span>¥${last.unit_price}/个</span>` : ''}</div>`;
    const existLog = purchaseTd.querySelector('.purchase-log');
    if (existLog) existLog.outerHTML = logHtml;
    else purchaseTd.insertAdjacentHTML('beforeend', logHtml);
  }
}

async function adjustStock(name, delta) {
  const part = allParts.find(p => p.name === name);
  if (!part) return;
  const newStock = Math.max(0, part.stock + delta);
  try {
    await fetch('/api/parts/' + encodeURIComponent(name) + '/stock', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ stock: newStock }),
    });
    part.stock = newStock;
    part.diff = newStock - part.demand;
    updatePartRowInPlace(part);
  } catch (e) { showToast('操作失败', 'danger'); }
}

async function doPurchase(name) {
  const qtyInput = document.getElementById('buy-' + escAttr(name));
  const priceInput = document.getElementById('price-' + escAttr(name));
  const qty = parseInt(qtyInput.value);
  const unit_price = parseFloat(priceInput.value) || null;
  if (!qty || qty <= 0) { showToast('请输入有效数量', 'warning'); return; }
  try {
    const res = await fetch('/api/parts/' + encodeURIComponent(name) + '/purchase', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ qty, unit_price }),
    });
    const data = await res.json();
    if (!data.ok) { showToast(data.error, 'warning'); return; }
    // 局部更新内存数据和 DOM，不重绘整表
    const part = allParts.find(p => p.name === name);
    if (part) {
      part.stock = data.stock;
      part.diff = data.stock - part.demand;
      part.purchases.push({
        qty, unit_price,
        date: new Date().toISOString().split('T')[0],
      });
      updatePartRowInPlace(part);
    }
    const priceStr = unit_price ? `，单价 ¥${unit_price}` : '';
    showToast(`${name} 入库 ${qty} 个${priceStr}`, 'success');
    qtyInput.value = '';
    priceInput.value = '';
  } catch (e) { showToast('操作失败', 'danger'); }
}

async function batchPurchase() {
  // 先收集所有待入库数据，避免 loadParts 重绘后 DOM 丢失
  const tasks = [];
  for (const p of allParts) {
    const qtyInput = document.getElementById('buy-' + escAttr(p.name));
    if (!qtyInput) continue;
    const qty = parseInt(qtyInput.value);
    if (!qty || qty <= 0) continue;
    const priceInput = document.getElementById('price-' + escAttr(p.name));
    const unit_price = parseFloat(priceInput ? priceInput.value : '') || null;
    tasks.push({ name: p.name, qty, unit_price });
  }
  if (tasks.length === 0) {
    showToast('没有填写数量的器件', 'warning');
    return;
  }
  // 先清空所有输入框，防止重复提交
  tasks.forEach(t => {
    const qi = document.getElementById('buy-' + escAttr(t.name));
    const pi = document.getElementById('price-' + escAttr(t.name));
    if (qi) qi.value = '';
    if (pi) pi.value = '';
  });
  let submitted = 0, failed = 0;
  for (const t of tasks) {
    try {
      const res = await fetch('/api/parts/' + encodeURIComponent(t.name) + '/purchase', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ qty: t.qty, unit_price: t.unit_price }),
      });
      const data = await res.json();
      if (data.ok) submitted++;
      else failed++;
    } catch (e) { failed++; }
  }
  if (failed > 0) {
    showToast(`入库完成：${submitted} 种成功，${failed} 种失败`, 'warning');
  } else {
    showToast(`已批量入库 ${submitted} 种器件`, 'success');
  }
  await loadParts();
}

async function syncParts() {
  try {
    const res = await fetch('/api/parts/sync', { method: 'POST' });
    const data = await res.json();
    showToast(`同步完成，新增 ${data.added} 个器件`, 'success');
    await loadParts();
  } catch (e) { showToast('同步失败', 'danger'); }
}

function setPartsFilter(btn, f) {
  document.querySelectorAll('[data-pf]').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  partsFilter = f;
  renderParts();
}

function filterParts() { renderParts(); }

function showLoading(show) {
  const loading = document.getElementById('loading');
  const wrap = document.getElementById('project-table-wrap');
  const empty = document.getElementById('empty-state');
  if (loading) loading.style.display = show ? 'block' : 'none';
  if (show) {
    if (wrap) wrap.classList.add('d-none');
    if (empty) empty.classList.add('d-none');
  }
}

function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  document.getElementById('toast-msg').textContent = msg;
  toast.className = `toast align-items-center border-0 text-bg-${type}`;
  bootstrap.Toast.getOrCreateInstance(toast, { delay: 2500 }).show();
}

// ─── 统计页 ───────────────────────────────────────────

let dailyStatsData = [];

async function loadDailyStats() {
  try {
    const res = await fetch('/api/stats/daily');
    dailyStatsData = await res.json();
    renderDailyStats();
  } catch (e) {
    showToast('统计数据加载失败：' + e.message, 'danger');
  }
}

function renderDailyStats() {
  const data = dailyStatsData.filter(d => d.date !== '未填日期');

  // 概览卡片
  if (data.length > 0) {
    const busiestDay = data.reduce((a, b) => a.count >= b.count ? a : b);
    const richestDay = data.reduce((a, b) => a.price >= b.price ? a : b);
    const totalDays = data.length;
    const totalOrders = data.reduce((s, d) => s + d.count, 0);
    const totalPrice = data.reduce((s, d) => s + d.price, 0);
    const avgPrice = totalOrders > 0 ? Math.round(totalPrice / totalOrders) : 0;

    document.getElementById('stats-busiest-day').textContent = busiestDay.date;
    document.getElementById('stats-busiest-day-count').textContent = `${busiestDay.count} 个单子`;
    document.getElementById('stats-richest-day').textContent = richestDay.date;
    document.getElementById('stats-richest-day-amount').textContent = `¥${formatMoney(richestDay.price)}`;
    document.getElementById('stats-total-days').textContent = totalDays;
    document.getElementById('stats-avg-price').textContent = `¥${formatMoney(avgPrice)}`;
  }

  // 明细表格（含未填日期）
  const tbody = document.getElementById('stats-tbody');
  const allData = dailyStatsData;

  if (allData.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-4">暂无数据</td></tr>`;
    return;
  }

  tbody.innerHTML = allData.map(d => {
    const statusColors = { '待绘制': 'secondary', '待打板': 'warning', '等板': 'info', '待交付': 'purple', '已完成': 'success' };
    const projTags = d.projects.map(p => {
      const color = statusColors[p.status] || 'secondary';
      const priceStr = p.price ? ` ¥${p.price}` : '';
      return `<span class="stats-proj-tag stats-proj-tag-${p.status}">${escHtml(p.name)}${priceStr}</span>`;
    }).join('');

    const isUnfilled = d.date === '未填日期';
    return `
<tr class="${isUnfilled ? 'stats-row-unfilled' : ''}">
  <td class="stats-date-cell">${isUnfilled ? '<span class="text-muted">未填日期</span>' : `<strong>${d.date}</strong>`}</td>
  <td style="text-align:center">
    <span class="stats-count-badge">${d.count}</span>
  </td>
  <td style="text-align:right;font-weight:600;color:#7c3aed">¥${formatMoney(d.price)}</td>
  <td style="text-align:right;font-weight:600;color:var(--success)">¥${formatMoney(d.prepay)}</td>
  <td><div class="stats-proj-tags">${projTags}</div></td>
</tr>`;
  }).join('');
}

// ─── 预报价 ───────────────────────────────────────────

async function loadEstimate() {
  if (!currentProject) return;
  const btn = document.querySelector('#estimate-wrap .btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>计算中...'; }

  try {
    const res = await fetch('/api/estimate/' + encodeURIComponent(currentProject.name));
    const data = await res.json();
    if (!data.ok) { showToast(data.error || '估算失败', 'danger'); return; }

    document.getElementById('estimate-result').style.display = 'block';
    document.getElementById('estimate-total').textContent = `¥${data.total_cost.toFixed(2)}`;

    // 缺少单价的器件提示
    const missingWrap = document.getElementById('estimate-missing-wrap');
    if (data.missing.length > 0) {
      missingWrap.style.display = 'block';
      document.getElementById('estimate-missing-text').textContent = data.missing.join('、');
    } else {
      missingWrap.style.display = 'none';
    }

    // 明细表格
    const tbody = document.getElementById('estimate-tbody');
    tbody.innerHTML = data.items.map(item => {
      const unitStr = item.unit_price !== null ? `¥${item.unit_price}` : '<span class="text-muted">—</span>';
      const subStr = item.subtotal !== null ? `¥${item.subtotal.toFixed(2)}` : '<span class="text-muted">—</span>';
      return `
<tr>
  <td>${escHtml(item.part)}</td>
  <td style="text-align:center">${item.qty}</td>
  <td style="text-align:right">${unitStr}</td>
  <td style="text-align:right;font-weight:600">${subStr}</td>
</tr>`;
    }).join('');
  } catch (e) {
    showToast('估算失败：' + e.message, 'danger');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-calculator me-1"></i>重新计算'; }
  }
}

// ─── 导出功能 ────────────────────────────────────────────

// 跟踪选中的文件列表
let selectedFilesForExport = [];

function renderFileListWithCheckboxes(files, projectName) {
  /**
   * 在侧边栏渲染文件列表，并添加复选框用于导出
   */
  const wrap = document.getElementById('sidebar-files');
  if (!wrap) return;

  selectedFilesForExport = [];  // 重置选择列表

  if (!files || files.length === 0) {
    wrap.innerHTML = '<p class="text-muted small">暂无文件</p>';
    updateExportButtonState();
    return;
  }

  wrap.innerHTML = files.map(file => `
    <div class="file-item">
      <input type="checkbox" class="file-checkbox" data-filename="${escHtml(file)}"
        onchange="updateExportSelection('${escHtml(file)}', this.checked)">
      <span class="file-name">${escHtml(file)}</span>
    </div>
  `).join('');

  updateExportButtonState();
}

function updateExportSelection(filename, checked) {
  /**
   * 更新选中的导出文件列表
   */
  if (checked) {
    if (!selectedFilesForExport.includes(filename)) {
      selectedFilesForExport.push(filename);
    }
  } else {
    selectedFilesForExport = selectedFilesForExport.filter(f => f !== filename);
  }
  updateExportButtonState();
}

// ─── 快速导出 PCB 制板文件（系统级功能） ───────────────
async function loadWaitingForPCBProjects() {
  /**
   * 加载所有处于「等板」阶段的项目并显示其PCB文件列表
   */
  const loadingEl = document.getElementById('quick-export-loading');
  const emptyEl = document.getElementById('quick-export-empty');
  const listEl = document.getElementById('quick-export-list');

  if (loadingEl) loadingEl.style.display = 'block';
  if (listEl) listEl.innerHTML = '';
  if (emptyEl) emptyEl.style.display = 'none';

  try {
    // 过滤出所有处于「等板」阶段的项目
    const waitingProjects = allProjects.filter(p => p.status === '等板');

    if (!waitingProjects || waitingProjects.length === 0) {
      if (loadingEl) loadingEl.style.display = 'none';
      if (emptyEl) emptyEl.style.display = 'block';
      return;
    }

    // 构建HTML
    let html = '';
    for (const project of waitingProjects) {
      // 获取项目的PCB文件
      const pcbFiles = (project.files || []).filter(f => f.includes('+PCB制板文件.zip'));
      const fileCount = pcbFiles.length;

      html += `
        <div class="quick-export-card">
          <div class="quick-export-card-header">
            <div class="quick-export-project-info">
              <h6 class="mb-1">${escHtml(project.name)}</h6>
              <small class="text-muted">${fileCount} 个PCB文件</small>
            </div>
            <button class="btn btn-primary btn-sm" onclick="quickExportProjectPCB('${escHtml(project.name)}')">
              <i class="bi bi-lightning-charge me-1"></i>导出
            </button>
          </div>
          <div class="quick-export-files">
            ${pcbFiles.length > 0
              ? pcbFiles.map(f => `<small class="file-item"><i class="bi bi-file-zip me-1"></i>${escHtml(f.split('/').pop())}</small>`).join('')
              : '<small class="text-muted">暂无PCB文件</small>'
            }
          </div>
        </div>
      `;
    }

    if (listEl) {
      listEl.innerHTML = html;
    }

  } catch (e) {
    showToast('加载失败：' + e.message, 'danger');
  } finally {
    if (loadingEl) loadingEl.style.display = 'none';
  }
}

async function quickExportProjectPCB(projectName) {
  /**
   * 快速导出单个项目的所有PCB制板文件
   */
  try {
    const project = allProjects.find(p => p.name === projectName);
    if (!project) {
      showToast('项目不存在', 'warning');
      return;
    }

    // 获取所有PCB文件
    const pcbFiles = (project.files || []).filter(f => f.includes('+PCB制板文件.zip'));
    if (pcbFiles.length === 0) {
      showToast('该项目没有PCB制板文件', 'warning');
      return;
    }

    // 发送导出请求
    const res = await fetch(`/api/projects/${encodeURIComponent(projectName)}/export-files`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: pcbFiles })
    });

    if (!res.ok) {
      throw new Error('导出失败');
    }

    // 下载ZIP文件
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // 从Content-Disposition获取文件名，或者自动生成
    const contentDisposition = res.headers.get('content-disposition');
    const filename = contentDisposition
      ? contentDisposition.split('filename=')[1]?.replace(/"/g, '')
      : `${projectName}_PCB文件_${new Date().toISOString().split('T')[0]}.zip`;

    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showToast(`成功导出 ${pcbFiles.length} 个PCB文件`, 'success');

    // 刷新导出历史
    if (currentPage === 'export-history') {
      loadExportHistory();
    }

  } catch (e) {
    showToast('导出失败：' + e.message, 'danger');
  }
}

function updateExportButtonState() {
  /**
   * 更新导出按钮的显示状态和计数
   */
  const exportSection = document.getElementById('export-section');
  const countSpan = document.getElementById('export-count');
  const quickExportBtn = document.getElementById('btn-quick-export-pcb');
  const projectStatus = document.getElementById('edit-status')?.value || '待绘制';
  const hasManualSelection = selectedFilesForExport.length > 0;
  const isWaitingForPCB = projectStatus === '等板';

  // 计算 export-section 应该显示的条件：
  // 1. 有手动选择的文件，或
  // 2. 项目在「等板」阶段
  const shouldShowExportSection = hasManualSelection || isWaitingForPCB;

  if (exportSection) {
    exportSection.style.display = shouldShowExportSection ? '' : 'none';
  }
  if (countSpan) {
    countSpan.textContent = selectedFilesForExport.length;
  }

  // 仅在「等板」阶段且没有手动选择时显示快速导出按钮
  if (quickExportBtn) {
    quickExportBtn.style.display = (isWaitingForPCB && !hasManualSelection) ? '' : 'none';
  }
}

function quickExportPCBFiles() {
  /**
   * 快速导出 - 自动选中所有 +PCB制板文件.zip 文件并导出
   */
  if (!currentProject) {
    showToast('请先打开项目', 'warning');
    return;
  }

  // 获取所有文件列表中的复选框
  const checkboxes = document.querySelectorAll('.file-checkbox');

  // 找到所有 PCB 制板文件
  let pcbFiles = [];
  checkboxes.forEach(checkbox => {
    const filename = checkbox.getAttribute('data-filename');
    if (filename && filename.includes('+PCB制板文件.zip')) {
      pcbFiles.push(filename);
      checkbox.checked = true;  // 自动勾选
    }
  });

  if (pcbFiles.length === 0) {
    showToast('当前项目没有 PCB 制板文件', 'warning');
    return;
  }

  // 更新选中的文件列表
  selectedFilesForExport = pcbFiles;
  updateExportButtonState();

  // 立即导出
  showToast(`已自动选中 ${pcbFiles.length} 个 PCB 制板文件，正在导出...`, 'info');
  setTimeout(() => {
    exportSelectedFiles();
  }, 500);
}

async function exportSelectedFiles() {
  /**
   * 导出选中的文件为 ZIP
   */
  if (!currentProject || selectedFilesForExport.length === 0) {
    showToast('请先选择要导出的文件', 'warning');
    return;
  }

  const btn = document.getElementById('btn-export-files');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>导出中...'; }

  try {
    const res = await fetch(`/api/projects/${encodeURIComponent(currentProject)}/export-files`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: selectedFilesForExport })
    });

    if (!res.ok) {
      const errData = await res.json();
      throw new Error(errData.error || '导出失败');
    }

    // 下载文件
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = res.headers.get('content-disposition')?.split('filename=')[1] || 'export.zip';
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showToast('文件已导出并下载', 'success');
    selectedFilesForExport = [];
    updateExportButtonState();

    // 刷新导出历史
    setTimeout(() => { if (currentPage === 'export-history') loadExportHistory(); }, 1000);
  } catch (e) {
    showToast('导出失败：' + e.message, 'danger');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-download me-1"></i>导出为ZIP'; }
  }
}

async function loadExportHistory(filterByType = null) {
  /**
   * 加载并显示导出历史
   * filterByType: 'pcb' | 'docx' | null - 如果指定，则按类型筛选并标注历史
   */
  const loading = document.getElementById('export-history-loading');
  const empty = document.getElementById('export-history-empty');
  const list = document.getElementById('export-history-list');

  // 更新筛选按钮状态
  document.querySelectorAll('.export-history-filter .filter-btn').forEach((btn, idx) => {
    btn.classList.remove('active');
    // 全部按钮 (索引0), PCB按钮 (索引1), DOCX按钮 (索引2)
    const btnType = idx === 0 ? null : (idx === 1 ? 'pcb' : 'docx');
    if (btnType === filterByType) {
      btn.classList.add('active');
    }
  });

  if (loading) loading.style.display = '';

  try {
    const res = await fetch('/api/export-history');
    const data = await res.json();

    if (!data.ok) throw new Error('加载失败');

    let records = data.records || [];

    if (loading) loading.style.display = 'none';

    if (records.length === 0) {
      if (empty) empty.style.display = '';
      if (list) list.innerHTML = '';
      return;
    }

    if (empty) empty.style.display = 'none';

    // 如果指定了导出类型，则按类型统计并排序
    let sortedRecords = records;
    if (filterByType) {
      // 分离已导出和未导出
      const exported = [];
      const notExported = [];

      records.forEach(record => {
        // 检查这条记录是否是该类型的导出
        const isThisType = record.export_type === filterByType;
        if (isThisType) {
          exported.push(record);
        } else {
          notExported.push(record);
        }
      });

      // 已导出的放在前面，未导出的放在下面
      sortedRecords = [...exported, ...notExported];
    }

    // 渲染导出记录
    if (list) {
      list.innerHTML = sortedRecords.map((record, index) => {
        // 判断这条记录是否标注
        const isMarked = filterByType && record.export_type === filterByType;
        const markClass = isMarked ? 'export-record-marked' : '';
        const markBadge = isMarked
          ? `<span class="export-type-badge" style="margin-left:8px;">${record.export_type === 'pcb' ? 'PCB' : 'DOCX'}</span>`
          : '';

        // 渲染项目列表（支持旧格式和新格式）
        let projectsHtml = '';
        if (record.projects && Array.isArray(record.projects)) {
          // 新格式：直接使用 projects 数组
          projectsHtml = record.projects.join('、');
        } else if (record.project_name) {
          // 旧格式：使用单个 project_name
          projectsHtml = record.project_name;
        }

        return `
          <div class="export-record-card ${markClass}">
            <div class="record-header">
              <div class="record-info">
                <h6 class="record-project">
                  ${escHtml(projectsHtml)}
                  ${markBadge}
                </h6>
                <small class="record-time">
                  <i class="bi bi-calendar-event me-1"></i>${record.export_time}
                </small>
              </div>
              <button class="btn btn-sm btn-outline-danger" onclick="deleteExportRecord('${record.id}')">
                <i class="bi bi-trash me-1"></i>删除
              </button>
            </div>
            ${record.files ? `
              <div class="record-files">
                <small class="text-muted">文件（${record.file_count}个）：</small>
                <ul class="file-list-compact">
                  ${record.files.slice(0, 5).map(f => '<li>' + escHtml(f) + '</li>').join('')}
                  ${record.file_count > 5 ? '<li class="text-muted">... 等 ' + (record.file_count - 5) + ' 个</li>' : ''}
                </ul>
              </div>
            ` : ''}
          </div>
        `;
      }).join('');
    }
  } catch (e) {
    showToast('加载导出历史失败：' + e.message, 'danger');
    if (loading) loading.style.display = 'none';
  }
}

async function deleteExportRecord(recordId) {
  /**
   * 删除一条导出记录
   */
  if (!confirm('确定要删除这条记录吗？')) return;

  try {
    const res = await fetch(`/api/export-history/${recordId}`, { method: 'DELETE' });
    const data = await res.json();

    if (!data.ok) throw new Error(data.error || '删除失败');

    showToast('记录已删除', 'success');
    loadExportHistory();
  } catch (e) {
    showToast('删除失败：' + e.message, 'danger');
  }
}

async function clearAllExportHistory() {
  /**
   * 清空所有导出记录
   */
  if (!confirm('确定要清空所有导出记录吗？此操作无法撤销。')) return;

  try {
    const res = await fetch('/api/export-history');
    const data = await res.json();
    const records = data.records || [];

    // 逐条删除所有记录
    for (const record of records) {
      await fetch(`/api/export-history/${record.id}`, { method: 'DELETE' });
    }

    showToast('已清空所有导出记录', 'success');
    loadExportHistory();
  } catch (e) {
    showToast('清空失败：' + e.message, 'danger');
  }
}

// ─── 窗口控制（pywebview 无边框模式）────────────────────
function wmMinimize() {
  if (window.pywebview) pywebview.api.minimize();
}
function wmClose() {
  if (window.pywebview) pywebview.api.close_window();
}

// ─── 通用导出系统 ─────────────────────────────────────────

let universalExportState = {
  selectedStages: [],
  selectedExportType: 'pcb',
  selectedProjects: [],
  filteredProjects: []
};

function initUniversalExport() {
  /**
   * 初始化通用导出系统 - 清空状态并准备UI
   */
  // 重置导出状态
  universalExportState.selectedStages = [];
  universalExportState.selectedProjects = [];
  universalExportState.selectedExportType = 'pcb';

  // 清空工具栏中的阶段复选框
  document.querySelectorAll('#export-filter-block input[name="stage"]').forEach(cb => {
    cb.checked = false;
  });

  // 默认加载所有项目（不按阶段筛选）
  loadUniversalExportProjects();
  updateUniversalExportUI();
}

function onStageFilterChange() {
  /**
   * 阶段筛选变化处理 - 重新加载项目列表
   */
  // 获取所有选中的阶段
  const checkboxes = document.querySelectorAll('input[name="stage"]:checked');
  universalExportState.selectedStages = Array.from(checkboxes).map(cb => cb.value);

  // 重置项目选择
  universalExportState.selectedProjects = [];

  // 加载项目列表
  loadUniversalExportProjects();

  // 更新UI
  updateUniversalExportUI();
}

function onExportTypeChange(type) {
  /**
   * 导出类型选择变化处理
   */
  universalExportState.selectedExportType = type;

  // 重置项目选择
  universalExportState.selectedProjects = [];

  // 重新加载项目列表（因为不同导出类型的条件不同）
  loadUniversalExportProjects();

  // 更新UI
  updateUniversalExportUI();
}

function loadUniversalExportProjects() {
  /**
   * 根据阶段和导出类型加载项目列表
   * 并标注已导出过的项目
   */
  const loadingEl = document.getElementById('universal-export-loading');
  const emptyEl = document.getElementById('universal-export-empty');
  const listEl = document.getElementById('universal-export-list');

  if (loadingEl) loadingEl.style.display = 'block';
  if (listEl) listEl.innerHTML = '';
  if (emptyEl) emptyEl.style.display = 'none';

  // 获取导出历史中已导出过的项目
  const exportedProjects = {};
  fetch('/api/export-history')
    .then(res => res.json())
    .then(data => {
      if (data.ok && data.records) {
        data.records.forEach(record => {
          const exportType = record.export_type || 'unknown';
          const projects = record.projects || [];

          projects.forEach(projectName => {
            if (!exportedProjects[projectName]) {
              exportedProjects[projectName] = {};
            }
            exportedProjects[projectName][exportType] = true;
          });
        });
      }
      // 导出历史加载完成后，重新渲染项目列表
      renderUniversalExportProjects(exportedProjects);
    })
    .catch(() => {
      // 如果获取历史失败，仍然继续加载项目列表（不标注）
      renderUniversalExportProjects({});
    });

  function renderUniversalExportProjects(exportedProjects) {
    try {
      // 筛选项目
      let filtered = allProjects;

    // 按阶段筛选
    if (universalExportState.selectedStages.length > 0) {
      filtered = filtered.filter(p => {
        const status = p.status || '待绘制';
        return universalExportState.selectedStages.includes(status);
      });
    }

    // 按导出类型筛选
    const exportType = universalExportState.selectedExportType;
    if (exportType === 'pcb') {
      // PCB导出：只显示包含PCB文件的项目
      filtered = filtered.filter(p => {
        const files = p.files || [];
        return files.some(f => f.includes('+PCB制板文件.zip'));
      });
    } else if (exportType === 'docx') {
      // DOCX导出：显示所有有元器件和功能的项目
      filtered = filtered.filter(p => {
        return (p.components || '').trim() !== '' && (p.functions || '').trim() !== '';
      });
    }

    universalExportState.filteredProjects = filtered;

    if (filtered.length === 0) {
      if (loadingEl) loadingEl.style.display = 'none';
      if (emptyEl) emptyEl.style.display = 'block';
      return;
    }

    // 渲染项目列表
    let html = '';

    // 全选复选框
    html += `
      <div class="universal-export-select-all">
        <label>
          <input type="checkbox" onchange="toggleSelectAllProjects(this.checked)">
          <span>全选</span>
          <span class="ml-2 text-muted">(共 ${filtered.length} 个项目)</span>
        </label>
      </div>
    `;

      // 分离未导出和已导出的项目
      const notExported = [];
      const exported = [];

      filtered.forEach(project => {
        const hasExported = exportedProjects[project.name] && exportedProjects[project.name][exportType];
        if (hasExported) {
          exported.push(project);
        } else {
          notExported.push(project);
        }
      });

      // 未导出的项目放前面，已导出的排在后面
      const sortedProjects = [...notExported, ...exported];

      // 项目列表
      html += '<div class="universal-export-projects">';
      sortedProjects.forEach(project => {
        const isSelected = universalExportState.selectedProjects.includes(project.name);
        const fileCount = exportType === 'pcb'
          ? (project.files || []).filter(f => f.includes('+PCB制板文件.zip')).length
          : 0;
        const fileInfo = exportType === 'pcb' ? `⭐ ${fileCount}个PCB文件` : '';

        // 检查是否已导出
        const isExported = exportedProjects[project.name] && exportedProjects[project.name][exportType];
        const exportedBadge = isExported
          ? `<span class="export-status-badge" title="已在此导出类型中导出过"><i class="bi bi-check-circle-fill"></i>已导出</span>`
          : '';
        const itemClass = isExported ? 'exported-project' : '';

        // 项目编号
        const idBadge = project.id
          ? `<span class="project-id-badge">[${String(project.id).padStart(3, '0')}]</span>`
          : '';

        html += `
          <div class="universal-export-project-item ${itemClass}">
            <label class="project-checkbox-label">
              <input type="checkbox" ${isSelected ? 'checked' : ''}
                onchange="toggleProjectSelection('${escHtml(project.name)}', this.checked)">
              ${idBadge}
              <span class="project-name">${escHtml(project.name)}</span>
              ${fileInfo ? `<span class="project-info">${fileInfo}</span>` : ''}
              ${exportedBadge}
            </label>
          </div>
        `;
      });
      html += '</div>';

      if (listEl) listEl.innerHTML = html;

    } catch (e) {
      showToast('加载失败：' + e.message, 'danger');
    } finally {
      if (loadingEl) loadingEl.style.display = 'none';
    }
  }
}

function toggleSelectAllProjects(checked) {
  /**
   * 全选/取消全选
   */
  if (checked) {
    universalExportState.selectedProjects = universalExportState.filteredProjects.map(p => p.name);
  } else {
    universalExportState.selectedProjects = [];
  }

  // 更新复选框
  document.querySelectorAll('.universal-export-project-item input[type="checkbox"]').forEach(cb => {
    cb.checked = checked;
  });

  updateUniversalExportUI();
}

function toggleProjectSelection(projectName, checked) {
  /**
   * 切换单个项目的选择状态
   */
  if (checked) {
    if (!universalExportState.selectedProjects.includes(projectName)) {
      universalExportState.selectedProjects.push(projectName);
    }
  } else {
    universalExportState.selectedProjects = universalExportState.selectedProjects.filter(p => p !== projectName);
  }

  updateUniversalExportUI();
}

function updateUniversalExportUI() {
  /**
   * 更新UI显示 - 计数和按钮状态
   */
  // 更新已选阶段文本
  const stageTexts = universalExportState.selectedStages.length > 0
    ? universalExportState.selectedStages.join('、')
    : '未选择';
  const selectedStagesEl = document.getElementById('selected-stages-text');
  if (selectedStagesEl) selectedStagesEl.textContent = stageTexts;

  // 更新项目计数
  const projectCountEl = document.getElementById('filtered-project-count');
  if (projectCountEl) projectCountEl.textContent = universalExportState.filteredProjects.length;

  // 更新选中项目计数
  const selectedCountEl = document.getElementById('selected-projects-count');
  if (selectedCountEl) selectedCountEl.textContent = universalExportState.selectedProjects.length;

  // 更新PCB/DOCX计数
  const exportType = universalExportState.selectedExportType;
  if (exportType === 'pcb') {
    const pcbCount = universalExportState.filteredProjects.filter(p => {
      const files = p.files || [];
      return files.some(f => f.includes('+PCB制板文件.zip'));
    }).length;
    const pcbCountEl = document.getElementById('pcb-count');
    if (pcbCountEl) pcbCountEl.textContent = pcbCount;
  } else if (exportType === 'docx') {
    const docxCount = universalExportState.filteredProjects.length;
    const docxCountEl = document.getElementById('docx-count');
    if (docxCountEl) docxCountEl.textContent = docxCount;
  }

  // 更新导出按钮状态
  const exportBtn = document.getElementById('btn-universal-export');
  if (exportBtn) {
    exportBtn.disabled = universalExportState.selectedProjects.length === 0;
  }
}

async function exportUniversal() {
  /**
   * 执行通用导出
   */
  if (universalExportState.selectedProjects.length === 0) {
    showToast('请选择要导出的项目', 'warning');
    return;
  }

  const btn = document.getElementById('btn-universal-export');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>导出中...';
  }

  try {
    const payload = {
      projects: universalExportState.selectedProjects,
      type: universalExportState.selectedExportType,
      stages: universalExportState.selectedStages
    };

    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      const errData = await res.json();
      throw new Error(errData.error || '导出失败');
    }

    // 获取导出的文件
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // 从Content-Disposition获取文件名
    const contentDisposition = res.headers.get('content-disposition');
    const filename = contentDisposition
      ? contentDisposition.split('filename=')[1]?.replace(/"/g, '')
      : `export_${new Date().toISOString().split('T')[0]}.${universalExportState.selectedExportType === 'pcb' ? 'zip' : 'docx'}`;

    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showToast(`成功导出 ${universalExportState.selectedProjects.length} 个项目`, 'success');

  } catch (e) {
    showToast('导出失败：' + e.message, 'danger');
  } finally {
    if (btn) {
      btn.disabled = universalExportState.selectedProjects.length === 0;
      btn.innerHTML = `<i class="bi bi-download me-2"></i>导出 (<span id="selected-projects-count">${universalExportState.selectedProjects.length}</span> 个项目)`;
    }
  }
}
