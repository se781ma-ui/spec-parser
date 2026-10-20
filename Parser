<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Парсер спецификаций</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<style>
:root {
  --bg:#eef2f7; --card:#ffffff; --card-alt:#f7f9fc; --border:#e0e7ef; --border-alt:#e2e8f0;
  --text:#1a2634; --text-sub:#5e6f8d; --text-mute:#7a8aa3;
  --primary:#1a73e8; --primary-hover:#1557b0; --primary-disabled:#a0b8d0;
  --log-bg:#13202b; --log-text:#dce6f0; --tag:#2b7be4;
  --green:#27ae60; --red:#a52a2a; --shadow:0 8px 35px rgba(0,0,0,.1);
}
body.dark {
  --bg:#14181f; --card:#1f242c; --card-alt:#262c36; --border:#2e3540; --border-alt:#2e3540;
  --text:#e6ecf3; --text-sub:#9aa7b8; --text-mute:#7a8aa3;
  --primary:#3b82f6; --primary-hover:#2563eb; --primary-disabled:#3a4a5c;
  --log-bg:#0a0d12; --log-text:#b8c6d8; --tag:#3b82f6;
  --green:#27ae60; --red:#e57373; --shadow:0 8px 35px rgba(0,0,0,.4);
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Segoe UI',sans-serif; background:var(--bg); display:flex; justify-content:center; align-items:flex-start; min-height:100vh; padding:20px; transition:background .25s,color .25s; color:var(--text); }
.container { background:var(--card); max-width:1500px; width:100%; border-radius:16px; padding:24px 32px; box-shadow:var(--shadow); border:1px solid var(--border); }
.topbar { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-bottom:14px; }
h1 { font-size:22px; font-weight:700; color:var(--text); }
.theme-toggle { background:var(--card-alt); border:1px solid var(--border-alt); border-radius:8px; padding:8px 14px; cursor:pointer; font-size:16px; color:var(--text); }
.theme-toggle:hover { background:var(--border); }
.tabs { display:flex; gap:6px; margin-bottom:20px; border-bottom:2px solid var(--border-alt); }
.tab { background:transparent; border:none; padding:10px 20px; font-size:14px; font-weight:600; color:var(--text-sub); cursor:pointer; border-bottom:3px solid transparent; }
.tab:hover { color:var(--primary); }
.tab.active { color:var(--primary); border-bottom-color:var(--primary); }
.tab-content { display:none; }
.tab-content.active { display:block; }
.block { background:var(--card-alt); border-radius:12px; padding:16px 20px 18px; margin-bottom:14px; border:1px solid var(--border-alt); }
.block-title { font-weight:600; font-size:14px; color:var(--text); margin-bottom:10px; display:flex; align-items:center; gap:8px; }
.block-title .tag { background:var(--primary); color:#fff; font-size:10px; border-radius:30px; padding:0 12px; line-height:20px; }
.dropzone { border:2px dashed var(--text-mute); border-radius:10px; padding:18px 14px; text-align:center; background:var(--card); cursor:pointer; }
.dropzone:hover, .dropzone.dragover { border-color:var(--primary); background:var(--card-alt); }
.dropzone .icon { font-size:32px; display:block; margin-bottom:4px; }
.dropzone .main-text { font-size:14px; font-weight:500; color:var(--text); }
.dropzone .hint { font-size:12px; color:var(--text-mute); margin-top:3px; }
.file-info { display:none; align-items:center; background:#e3f0e5; padding:8px 14px; border-radius:8px; margin-top:10px; font-size:13px; color:#1a5a2a; gap:8px; flex-wrap:wrap; }
body.dark .file-info { background:#1e3a24; color:#a8e3b8; }
.file-info.show { display:flex; }
.file-info .name { font-weight:500; word-break:break-all; }
.file-info .size { color:#3d6b4a; font-size:12px; }
body.dark .file-info .size { color:#7fae90; }
.file-info .remove { margin-left:auto; background:none; border:none; color:var(--red); font-size:18px; font-weight:700; cursor:pointer; padding:0 6px; }
.btn-primary { border:none; border-radius:10px; font-size:16px; font-weight:600; cursor:pointer; width:100%; padding:14px 0; background:var(--primary); color:#fff; transition:.15s; margin:4px 0 12px; }
.btn-primary:hover { background:var(--primary-hover); }
.btn-primary:disabled { background:var(--primary-disabled); cursor:not-allowed; }
.btn-secondary { border:1px solid var(--border-alt); border-radius:10px; font-size:13px; font-weight:500; cursor:pointer; padding:8px 16px; background:var(--card-alt); color:var(--text); }
.btn-secondary:hover { background:var(--border); }
.btn-secondary:disabled { opacity:.4; cursor:not-allowed; }
.log-box { background:var(--log-bg); color:var(--log-text); border-radius:10px; padding:12px 16px; font-family:Consolas,monospace; font-size:12px; min-height:100px; max-height:220px; overflow-y:auto; line-height:1.6; white-space:pre-wrap; word-break:break-word; margin-top:10px; }
.log-box .ok { color:#7ddf9a; } .log-box .err { color:#f7a1a1; } .log-box .inf { color:#7bc9f0; } .log-box .warn { color:#ffd54f; }
.sub-files-list { margin-top:10px; font-size:13px; }
.sub-file-item { display:flex; align-items:center; gap:8px; background:var(--card); padding:5px 12px; border-radius:6px; margin-bottom:5px; border:1px solid var(--border-alt); }
.sub-file-item .remove-sub { margin-left:auto; background:none; border:none; color:var(--red); font-size:16px; cursor:pointer; }

.modal-overlay { position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,.5); display:none; justify-content:center; align-items:center; z-index:1000; }
.modal-overlay.show { display:flex; }
.modal-content { background:var(--card); color:var(--text); border-radius:16px; padding:22px; max-width:1500px; width:96%; max-height:92vh; overflow-y:auto; box-shadow:0 10px 40px rgba(0,0,0,.3); }
.modal-small { max-width:520px; }
.modal-header { font-size:18px; font-weight:600; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; }
.modal-footer { display:flex; justify-content:space-between; gap:10px; margin-top:16px; flex-wrap:wrap; }
.order-tree-picker { display:flex; align-items:center; gap:6px; font-size:13px; font-weight:500; }
.order-tree-picker select { padding:6px 10px; border:1px solid var(--border-alt); border-radius:6px; font-size:13px; background:var(--card); color:var(--text); max-width:280px; }

.preview-config { background:var(--card-alt); border:1px solid var(--border-alt); border-radius:10px; margin-bottom:12px; }
.preview-config-header { padding:8px 14px; cursor:pointer; font-weight:600; font-size:13px; display:flex; align-items:center; gap:8px; }
.preview-config-body { display:none; padding:12px 14px; border-top:1px solid var(--border-alt); }
.preview-config.open .preview-config-body { display:block; }

.preview-list { font-size:13px; }
.preview-header-row { display:flex; align-items:center; gap:6px; padding:6px 0; border-bottom:2px solid var(--border-alt); font-weight:700; font-size:11px; color:var(--text-sub); text-transform:uppercase; position:sticky; top:0; background:var(--card); z-index:2; }
.preview-header-row .h-del { width:24px; }
.preview-header-row .h-cb { width:18px; }
.preview-header-row .h-dot { width:16px; }
.preview-header-row .h-num { min-width:75px; }
.preview-header-row .h-name { flex:2; }
.preview-header-row .h-des { flex:1.5; }
.preview-header-row .h-local { width:70px; }
.preview-header-row .h-total { width:70px; }
.preview-header-row .h-note { width:180px; }
.preview-item { display:flex; align-items:center; gap:6px; padding:3px 0; border-bottom:1px solid var(--border-alt); min-height:30px; }
.preview-item .item-del { width:24px; flex-shrink:0; text-align:center; }
.preview-item .item-del button { background:none; border:none; color:var(--red); font-size:14px; cursor:pointer; padding:0 4px; opacity:.7; }
.preview-item .item-number { min-width:75px; color:var(--text-sub); font-family:monospace; white-space:pre; flex-shrink:0; font-size:12px; }
.preview-item .item-name { flex:2; }
.preview-item .item-des { flex:1.5; }
.preview-item .item-local { width:70px; flex-shrink:0; }
.preview-item .item-total { width:70px; flex-shrink:0; color:var(--green); font-weight:600; text-align:right; padding-right:6px; }
.preview-item .item-note { width:180px; flex-shrink:0; }
.preview-item input.edit-field { border:1px solid var(--border-alt); border-radius:4px; padding:2px 6px; font-size:12px; width:100%; background:var(--card); color:var(--text); }
.preview-item input.edit-field:focus { outline:none; border-color:var(--primary); }
.preview-item .expand-dot { display:inline-block; width:14px; height:14px; border-radius:50%; border:2px solid var(--text-sub); flex-shrink:0; cursor:pointer; }
.preview-item .expand-dot.filled { background:var(--text-sub); }
.preview-item .expand-dot.empty { background:var(--card); }
.preview-item .expand-dot.disabled { opacity:.25; cursor:default; }
.add-row-line { display:flex; align-items:center; gap:6px; padding:4px 0; }
.add-row-line .item-del { width:24px; flex-shrink:0; }
.add-row-line .add-btn { background:#e8f5ec; border:1px dashed var(--green); color:var(--green); font-weight:700; font-size:13px; padding:3px 14px; border-radius:6px; cursor:pointer; }
body.dark .add-row-line .add-btn { background:#1a3020; }
.tree-separator { background:var(--card-alt); padding:8px 14px; border-radius:6px; margin:14px 0 6px; font-weight:700; color:var(--text-sub); font-size:13px; }

.config-page-header { display:flex; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:16px; }
.config-page-header select { padding:8px 14px; border:1px solid var(--border-alt); border-radius:8px; font-size:14px; background:var(--card); color:var(--text); min-width:280px; }
.config-actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:12px; }
.config-grid { display:grid; grid-template-columns:repeat(2,1fr); gap:12px 20px; }
.config-grid .field { display:flex; flex-direction:column; gap:4px; }
.config-grid .field label { font-size:12px; color:var(--text-sub); font-weight:500; }
.config-grid .field input, .config-grid .field select { padding:8px 12px; border:1px solid var(--border-alt); border-radius:6px; font-size:13px; background:var(--card); color:var(--text); }
.config-grid .field input:focus, .config-grid .field select:focus { outline:none; border-color:var(--primary); }
.config-list { display:flex; flex-direction:column; gap:6px; margin-top:12px; }
.config-list .item { display:flex; align-items:center; gap:10px; background:var(--card); border:1px solid var(--border-alt); padding:8px 14px; border-radius:8px; font-size:13px; }
.config-list .item.active { border-color:var(--primary); background:var(--card-alt); }
.config-list .item .name { flex:1; font-weight:500; }
.config-list .item .meta { font-size:11px; color:var(--text-mute); }
.config-list .item .remove { background:none; border:none; color:var(--red); font-size:14px; cursor:pointer; padding:0 6px; }

.preview-config .config-grid { grid-template-columns:repeat(2,1fr); }
.preview-config-full { grid-column:1 / -1; }

@media (max-width:600px) { .container { padding:16px; } h1 { font-size:18px; } .config-grid { grid-template-columns:1fr; } }
</style>
</head>
<body>
<div class="container">
  <div class="topbar">
    <h1>📄 Парсер спецификаций</h1>
    <button class="theme-toggle" id="themeToggle" title="Тема">🌙</button>
  </div>

  <div class="tabs">
    <button class="tab active" data-tab="parser">📄 Парсер</button>
    <button class="tab" data-tab="configs">⚙️ Конфиги</button>
  </div>

  <div class="tab-content active" id="tab-parser">
    <div class="block">
      <div class="block-title"><span>📂 1. Главная спецификация</span><span class="tag">.xlsx / .docx</span></div>
      <div class="dropzone" id="mainDropZone">
        <span class="icon">📄</span>
        <div class="main-text">Нажмите или перетащите главный файл</div>
        <div class="hint">Поддерживаются .xlsx и .docx</div>
        <input type="file" id="mainFileInput" accept=".xlsx,.docx" style="display:none;">
      </div>
      <div class="file-info" id="mainFileInfo">
        <span>📎</span><span class="name" id="mainFileName"></span><span class="size" id="mainFileSize"></span>
        <button class="remove" id="removeMainFile">✕</button>
      </div>
    </div>

    <div class="block">
      <div class="block-title"><span>📁 2. Подсборки</span><span class="tag">.xlsx / .docx (можно несколько)</span></div>
      <div class="dropzone" id="subDropZone">
        <span class="icon">📚</span>
        <div class="main-text">Перетащите файлы подсборок (в любом порядке)</div>
        <div class="hint">Раскроются после нажатия «Построить дерево»</div>
        <input type="file" id="subFileInput" accept=".xlsx,.docx" multiple style="display:none;">
      </div>
      <div class="sub-files-list" id="subFilesList"></div>
    </div>

    <button class="btn-primary" id="buildBtn" disabled>▶ ПОСТРОИТЬ ДЕРЕВО (предпросмотр)</button>
    <div class="log-box" id="logBox"></div>
  </div>

  <div class="tab-content" id="tab-configs">
    <div class="block">
      <div class="block-title">📋 Мои конфиги</div>
      <div class="config-page-header">
        <label style="font-size:13px;color:var(--text-sub);font-weight:500;">Активный:</label>
        <select id="configSelect"></select>
        <button class="btn-secondary" id="newConfigBtn">➕ Создать</button>
        <button class="btn-secondary" id="renameConfigBtn">✏️ Переименовать</button>
        <button class="btn-secondary" id="deleteConfigBtn">🗑 Удалить</button>
      </div>
      <div class="config-list" id="configList"></div>
    </div>

    <div class="block">
      <div class="block-title">📝 Поля конфига</div>
      <div class="config-grid" id="configFieldsGrid"></div>
      <div class="config-actions">
        <button class="btn-primary" style="width:auto;padding:10px 24px;margin:0;" id="saveConfigBtn">💾 Сохранить</button>
        <button class="btn-secondary" id="downloadConfigBtn">⬇️ Скачать JSON</button>
        <button class="btn-secondary" id="uploadConfigBtn">⬆️ Загрузить JSON</button>
        <input type="file" id="uploadConfigInput" accept=".json" style="display:none;">
      </div>
    </div>
  </div>
</div>

<div class="modal-overlay" id="previewModal">
  <div class="modal-content">
    <div class="modal-header">
      <span>📋 Предпросмотр дерева</span>
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <div class="order-tree-picker" id="orderTreePicker" style="display:none;">
          <label for="orderTreeSelect">Дерево для заявки:</label>
          <select id="orderTreeSelect"></select>
        </div>
        <button class="btn-secondary" id="expandAllBtn">📂 Раскрыть всё</button>
        <button class="btn-secondary" id="collapseAllBtn">📁 Скрыть всё</button>
      </div>
    </div>

    <div class="preview-config" id="previewConfig">
      <div class="preview-config-header" id="previewConfigHeader">
        ⚙️ Параметры заявки <span style="margin-left:auto;font-size:11px;color:var(--text-sub);">▼ развернуть</span>
      </div>
      <div class="preview-config-body">
        <div style="margin-bottom:12px;">
          <label style="font-size:12px;color:var(--text-sub);font-weight:500;display:block;margin-bottom:6px;">Конфиг:</label>
          <select id="previewConfigSelect" style="padding:8px 12px;border:1px solid var(--border-alt);border-radius:6px;font-size:13px;background:var(--card);color:var(--text);width:100%;"></select>
        </div>
        <div style="margin-bottom:12px;">
          <label style="font-size:12px;color:var(--text-sub);font-weight:500;display:block;margin-bottom:6px;">Срок (общий для заявки):</label>
          <input type="text" id="previewSrokInput" style="padding:8px 12px;border:1px solid var(--border-alt);border-radius:6px;font-size:13px;background:var(--card);color:var(--text);width:100%;">
        </div>
        <div class="config-grid" id="previewConfigGrid"></div>
      </div>
    </div>

    <div class="preview-list" id="previewList"></div>
    <div class="modal-footer">
      <button class="btn-secondary" id="closePreview">Закрыть</button>
      <button class="btn-primary" style="width:auto;padding:10px 24px;margin:0;" id="downloadReportBtn">📥 Скачать отчёт (XLSX)</button>
      <button class="btn-primary" style="width:auto;padding:10px 24px;margin:0;background:var(--green);" id="generateOrderBtn" disabled>📋 Сформировать заявку</button>
    </div>
  </div>
</div>

<div class="modal-overlay" id="orderConfigModal">
  <div class="modal-content modal-small">
    <div class="modal-header">📋 Сформировать заявку</div>
    <div style="margin-bottom:12px;">
      <label style="font-size:13px;color:var(--text-sub);font-weight:500;display:block;margin-bottom:6px;">Конфиг (шапка заявки):</label>
      <select id="orderConfigSelect" style="padding:8px 14px;border:1px solid var(--border-alt);border-radius:8px;font-size:14px;background:var(--card);color:var(--text);width:100%;"></select>
      <div style="font-size:11px;color:var(--text-mute);margin-top:6px;">Если не выбран — заявка с пустой шапкой.</div>
    </div>
    <div style="margin-bottom:12px;">
      <label style="font-size:13px;color:var(--text-sub);font-weight:500;display:block;margin-bottom:6px;">Срок:</label>
      <input type="text" id="orderSrokInput" style="padding:8px 14px;border:1px solid var(--border-alt);border-radius:8px;font-size:14px;background:var(--card);color:var(--text);width:100%;">
    </div>
    <div class="modal-footer">
      <button class="btn-secondary" id="closeOrderConfigModal">Отмена</button>
      <button class="btn-primary" style="width:auto;padding:10px 24px;margin:0;background:var(--green);" id="confirmGenerateOrderBtn">📋 Сформировать</button>
    </div>
  </div>
</div>

<script>
(function(){
'use strict';

const CONFIG_FIELDS = [
  { key:'izdelie',     label:'Изделие' },
  { key:'decimal',     label:'Децимальный номер изделия' },
  { key:'data',        label:'Дата' },
  { key:'dogovor',     label:'Договор' },
  { key:'ooo',         label:'ООО "VXI-Системы"' },
  { key:'priemka',     label:'Приемка' },
  { key:'garantiya',   label:'Гарантия' },
  { key:'postavshik',  label:'Поставщик' },
  { key:'zakazchik',   label:'Заказчик' },
  { key:'vedushiy',    label:'Ведущий' },
  { key:'konstruktor', label:'Конструктор' },
  { key:'inzhKonstr',  label:'Инженер-конструктор' },
  { key:'vedInzh',     label:'Ведущий инженер' },
  { key:'nachOtdela',  label:'Начальник отдела' }
];

const LS_KEY = 'parser_configs_v1';
const LS_ACTIVE = 'parser_active_config_v1';
const LS_THEME = 'parser_theme_v1';

const MARKER = 'Переменные данные для исполнений';
const KEEP_SECTIONS = ['Сборочные единицы', 'Детали', 'Жгуты'];
const IGNORE_SECTIONS = ['Документация', 'Стандартные изделия', 'Прочие изделия',
                         'Материалы', 'Комплекты', 'Программные изделия'];
const ALL_SECTIONS = KEEP_SECTIONS.concat(IGNORE_SECTIONS);
const BLOCK_HEADER_RE = /^[А-ЯЁA-Z]{2,}\.[\d.]+(-\d+)?$/i;
const SUFFIX_ONLY_RE = /^-\d+$/;
const DES_ANYWHERE_RE = /[А-ЯЁA-Z]{2,}\.[\d.]+(?:-\d+)?/gi;

let configs = {};
let activeConfigName = '';
let previewSelectedConfigName = '';
let orderSessionConfig = {};
let mainBuffer = null, mainName = '';
let subFiles = [];
let trees = [];
let flatNodes = [];

function $(id){ return document.getElementById(id); }
function escapeHtml(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function cleanText(t) {
  if (!t) return '';
  return String(t).replace(/\u00A0/g,' ').replace(/\r\n/g,' ').replace(/\n/g,' ')
    .replace(/\t/g,' ').replace(/\s+/g,' ').trim();
}
function log(msg, type) {
  const box = $('logBox');
  if (!box) return;
  const cls = type === 'ok' ? 'ok' : type === 'err' ? 'err' : type === 'warn' ? 'warn' : 'inf';
  const span = document.createElement('span');
  span.className = cls;
  span.textContent = msg + '\n';
  box.appendChild(span);
  box.scrollTop = box.scrollHeight;
}
function clearLog(){ const box=$('logBox'); if (box) box.innerHTML=''; }

function isStopWord(text) {
  if (!text) return false;
  return String(text).toLowerCase().includes('примечани');
}
function isIgnoredSheetName(name) {
  if (!name) return false;
  const n = name.trim().toLowerCase();
  if (n === 'лист регистрации изменений') return true;
  if (n.startsWith('листшаблон')) return true;
  return false;
}
function baseDesignation(des) {
  if (!des) return '';
  return String(des).trim().replace(/(СБ|СП|Э\d|ПЭ\d|ТУ|ФО|МЧ|Д|РЭ|ВЭ|ОП\d)$/i,'').trim();
}
function detectSection(name) {
  const n = String(name||'').toLowerCase();
  for (const s of ALL_SECTIONS) if (n.includes(s.toLowerCase())) return s;
  return null;
}
function emptyConfig() {
  const c = {};
  CONFIG_FIELDS.forEach(f => c[f.key] = '');
  return c;
}

// ТЕМА
function applyTheme(theme) {
  if (theme === 'dark') {
    document.body.classList.add('dark');
    $('themeToggle').textContent = '☀️';
  } else {
    document.body.classList.remove('dark');
    $('themeToggle').textContent = '🌙';
  }
  localStorage.setItem(LS_THEME, theme);
}
$('themeToggle').addEventListener('click', () => {
  const cur = document.body.classList.contains('dark') ? 'dark' : 'light';
  applyTheme(cur === 'dark' ? 'light' : 'dark');
});
applyTheme(localStorage.getItem(LS_THEME) || 'light');

// ВКЛАДКИ
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    tab.classList.add('active');
    $('tab-' + tab.dataset.tab).classList.add('active');
  });
});

// КОНФИГИ
function loadConfigsFromLS() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    configs = raw ? JSON.parse(raw) : {};
    if (typeof configs !== 'object' || Array.isArray(configs)) configs = {};
  } catch (e) { configs = {}; }
  activeConfigName = localStorage.getItem(LS_ACTIVE) || '';
  if (activeConfigName && !configs[activeConfigName]) activeConfigName = '';
  if (!activeConfigName) {
    const names = Object.keys(configs);
    if (names.length) activeConfigName = names[0];
  }
}
function saveConfigsToLS() {
  localStorage.setItem(LS_KEY, JSON.stringify(configs));
  localStorage.setItem(LS_ACTIVE, activeConfigName);
}
const configSelect = $('configSelect');
const configList = $('configList');
const configFieldsGrid = $('configFieldsGrid');
function renderConfigSelect() {
  configSelect.innerHTML = '';
  const names = Object.keys(configs);
  if (!names.length) {
    const opt = document.createElement('option');
    opt.value = ''; opt.textContent = '— нет конфигов —';
    configSelect.appendChild(opt);
  } else {
    names.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name; opt.textContent = name;
      if (name === activeConfigName) opt.selected = true;
      configSelect.appendChild(opt);
    });
  }
}
function renderConfigList() {
  configList.innerHTML = '';
  Object.keys(configs).forEach(name => {
    const item = document.createElement('div');
    item.className = 'item' + (name === activeConfigName ? ' active' : '');
    item.innerHTML = '<span class="name">' + escapeHtml(name) + '</span>' +
                     '<span class="meta">' + (name === activeConfigName ? '● активный' : '') + '</span>' +
                     '<button class="remove" data-name="' + escapeHtml(name) + '" title="Удалить">🗑</button>';
    configList.appendChild(item);
  });
  configList.querySelectorAll('.remove').forEach(b => {
    b.addEventListener('click', () => {
      const name = b.dataset.name;
      if (!confirm('Удалить конфиг "' + name + '"?')) return;
      delete configs[name];
      if (activeConfigName === name) activeConfigName = Object.keys(configs)[0] || '';
      saveConfigsToLS();
      refreshConfigUI();
    });
  });
}
function renderConfigFields() {
  const c = configs[activeConfigName] || emptyConfig();
  configFieldsGrid.innerHTML = '';
  CONFIG_FIELDS.forEach(f => {
    const div = document.createElement('div');
    div.className = 'field';
    div.innerHTML = '<label>' + escapeHtml(f.label) + '</label>' +
                    '<input type="text" data-key="' + f.key + '" value="' + escapeHtml(c[f.key]||'') + '">';
    configFieldsGrid.appendChild(div);
  });
}
function refreshConfigUI() {
  renderConfigSelect();
  renderConfigList();
  renderConfigFields();
  rebuildOrderConfigSelect();
  rebuildPreviewConfigSelect();
}
function rebuildOrderConfigSelect() {
  const sel = $('orderConfigSelect');
  sel.innerHTML = '';
  const none = document.createElement('option');
  none.value = ''; none.textContent = '— без конфига (пустая шапка) —';
  sel.appendChild(none);
  Object.keys(configs).forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    if (name === activeConfigName) opt.selected = true;
    sel.appendChild(opt);
  });
}
function rebuildPreviewConfigSelect() {
  const sel = $('previewConfigSelect');
  sel.innerHTML = '';
  const none = document.createElement('option');
  none.value = ''; none.textContent = '— не выбран —';
  sel.appendChild(none);
  Object.keys(configs).forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    sel.appendChild(opt);
  });
  if (previewSelectedConfigName && configs[previewSelectedConfigName]) {
    sel.value = previewSelectedConfigName;
  }
}
configSelect.addEventListener('change', () => {
  activeConfigName = configSelect.value;
  saveConfigsToLS();
  refreshConfigUI();
});
$('newConfigBtn').addEventListener('click', () => {
  const name = prompt('Название нового конфига:', 'Конфиг ' + (Object.keys(configs).length + 1));
  if (!name || !name.trim()) return;
  if (configs[name]) { alert('Такое имя уже есть'); return; }
  configs[name] = emptyConfig();
  activeConfigName = name;
  saveConfigsToLS();
  refreshConfigUI();
});
$('renameConfigBtn').addEventListener('click', () => {
  if (!activeConfigName) { alert('Нет активного конфига'); return; }
  const oldName = activeConfigName;
  const newName = prompt('Новое название:', oldName);
  if (!newName || !newName.trim() || newName === oldName) return;
  if (configs[newName]) { alert('Такое имя уже есть'); return; }
  configs[newName] = configs[oldName];
  delete configs[oldName];
  activeConfigName = newName;
  saveConfigsToLS();
  refreshConfigUI();
});
$('deleteConfigBtn').addEventListener('click', () => {
  if (!activeConfigName) return;
  if (!confirm('Удалить конфиг "' + activeConfigName + '"?')) return;
  delete configs[activeConfigName];
  activeConfigName = Object.keys(configs)[0] || '';
  saveConfigsToLS();
  refreshConfigUI();
});
$('saveConfigBtn').addEventListener('click', () => {
  if (!activeConfigName) { alert('Сначала создайте конфиг'); return; }
  const c = {};
  configFieldsGrid.querySelectorAll('input[data-key]').forEach(inp => {
    c[inp.dataset.key] = inp.value;
  });
  configs[activeConfigName] = c;
  saveConfigsToLS();
  refreshConfigUI();
  alert('Конфиг "' + activeConfigName + '" сохранён');
});
$('downloadConfigBtn').addEventListener('click', () => {
  if (!activeConfigName) { alert('Нет активного конфига'); return; }
  const c = {};
  configFieldsGrid.querySelectorAll('input[data-key]').forEach(inp => {
    c[inp.dataset.key] = inp.value;
  });
  configs[activeConfigName] = c;
  saveConfigsToLS();
  const blob = new Blob([JSON.stringify(c, null, 2)], { type:'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = activeConfigName.replace(/[\\/*?:"<>|]/g,'_') + '.json';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
});
$('uploadConfigBtn').addEventListener('click', () => $('uploadConfigInput').click());
$('uploadConfigInput').addEventListener('change', function() {
  if (!this.files.length) return;
  const file = this.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const obj = JSON.parse(e.target.result);
      let name = file.name.replace(/\.json$/i,'');
      if (configs[name]) name = name + '_' + Date.now();
      const c = emptyConfig();
      CONFIG_FIELDS.forEach(f => { if (obj[f.key] !== undefined) c[f.key] = String(obj[f.key]||''); });
      configs[name] = c;
      activeConfigName = name;
      saveConfigsToLS();
      refreshConfigUI();
      alert('Конфиг "' + name + '" загружен');
    } catch (err) { alert('Ошибка чтения JSON: ' + err.message); }
  };
  reader.readAsText(file, 'UTF-8');
  this.value = '';
});

// ПАРСЕР XLSX
async function parseXLSX(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbDoc = parser.parseFromString(wbXml, 'text/xml');
  const sheetEls = Array.from(wbDoc.querySelectorAll('sheet'));
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relsDoc = parser.parseFromString(relsXml, 'text/xml');
  const relsMap = {};
  for (const r of relsDoc.querySelectorAll('Relationship')) {
    relsMap[r.getAttribute('Id')] = r.getAttribute('Target');
  }
  const sharedStrings = [];
  const ssFile = zip.file('xl/sharedStrings.xml');
  if (ssFile) {
    const ssXml = await ssFile.async('string');
    const ssDoc = parser.parseFromString(ssXml, 'text/xml');
    ssDoc.querySelectorAll('si').forEach(si => {
      let text = '';
      si.querySelectorAll('t').forEach(t => text += t.textContent);
      sharedStrings.push(cleanText(text));
    });
  }
  function getCellValue(cell) {
    const t = cell.getAttribute('t');
    const v = cell.querySelector('v');
    const val = v ? v.textContent : '';
    if (t === 's') {
      const idx = parseInt(val,10);
      return (idx>=0 && idx<sharedStrings.length) ? sharedStrings[idx] : '';
    }
    if (t === 'inlineStr') {
      const tEl = cell.querySelector('is t');
      return tEl ? tEl.textContent : '';
    }
    return val;
  }
  const allRows = [];
  for (const sheetEl of sheetEls) {
    const sheetName = (sheetEl.getAttribute('name')||'').trim();
    if (isIgnoredSheetName(sheetName)) continue;
    const rid = sheetEl.getAttribute('r:id');
    const target = relsMap[rid];
    if (!target) continue;
    const sheetPath = 'xl/' + target.replace(/^\//,'');
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    const sheetXml = await sheetFile.async('string');
    const sheetDoc = parser.parseFromString(sheetXml, 'text/xml');
    const rowEls = sheetDoc.querySelectorAll('row');
    const sheetRows = [];
    let stop = false;
    rowEls.forEach(rowEl => {
      if (stop) return;
      const cellMap = {};
      rowEl.querySelectorAll('c').forEach(cell => {
        const ref = cell.getAttribute('r')||'';
        const m = ref.match(/^([A-Z]+)/);
        if (!m) return;
        const col = m[1];
        let idx = 0;
        for (let i=0;i<col.length;i++) idx = idx*26 + (col.charCodeAt(i)-65);
        cellMap[idx] = cleanText(getCellValue(cell));
      });
      let hasStop = false;
      for (const k in cellMap) {
        if (isStopWord(cellMap[k])) { hasStop = true; break; }
      }
      if (hasStop) { stop = true; return; }
      sheetRows.push(cellMap);
    });
    let headerRowIdx = -1;
    for (let i=0; i<Math.min(sheetRows.length, 20); i++) {
      const row = sheetRows[i];
      let hasDes = false, hasName = false;
      for (const key in row) {
        const val = (row[key]||'').toLowerCase();
        if (val.includes('обознач')) hasDes = true;
        if (val.includes('наимен') && !val.includes('докум')) hasName = true;
      }
      if (hasDes && hasName) { headerRowIdx = i; break; }
    }
    if (headerRowIdx < 0) continue;
    for (let i = headerRowIdx+1; i < sheetRows.length; i++) {
      const row = sheetRows[i];
      const des = row[5] || '';
      const name = row[6] || '';
      const count = row[7] || '';
      const note = row[8] || '';
      if (!des.trim() && !name.trim()) continue;
      if (name.trim().toLowerCase() === 'лист' && !des.trim()) continue;
      allRows.push({ des, name, count, note });
    }
  }
  return allRows;
}

// ПАРСЕР DOCX
async function parseDOCX(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();
  const xml = await zip.file('word/document.xml').async('string');
  const doc = parser.parseFromString(xml, 'text/xml');
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const tables = doc.getElementsByTagNameNS(W_NS, 'tbl');
  const allRows = [];
  let stop = false;
  for (const tbl of tables) {
    if (stop) break;
    const trs = tbl.getElementsByTagNameNS(W_NS, 'tr');
    for (const tr of trs) {
      if (stop) break;
      const cells = [];
      const tcs = tr.getElementsByTagNameNS(W_NS, 'tc');
      for (const tc of tcs) {
        let text = '';
        const ts = tc.getElementsByTagNameNS(W_NS, 't');
        for (const t of ts) text += t.textContent;
        const syms = tc.getElementsByTagNameNS(W_NS, 'sym');
        for (const sym of syms) {
          const ch = sym.getAttributeNS(W_NS, 'char');
          if (ch === 'F0D1') text += ' ▼';
        }
        const cleaned = cleanText(text);
        if (isStopWord(cleaned)) { stop = true; break; }
        cells.push(cleaned);
      }
      if (stop) break;
      const des = cells[3] || '';
      const name = cells[4] || '';
      const count = cells[5] || '';
      const note = cells[6] || '';
      if (!des.trim() && !name.trim()) continue;
      if (name.trim().toLowerCase() === 'лист' && !des.trim()) continue;
      allRows.push({ des, name, count, note });
    }
  }
  return allRows;
}

async function parseAnyFile(arrayBuffer, fileName) {
  const lower = (fileName || '').toLowerCase();
  if (lower.endsWith('.xlsx')) return await parseXLSX(arrayBuffer);
  if (lower.endsWith('.docx')) return await parseDOCX(arrayBuffer);
  throw new Error('Неподдерживаемый формат: ' + fileName);
}

// ДЕРЕВЬЯ
function buildTreesFromRows(rows) {
  let markerIdx = -1;
  for (let i=0;i<rows.length;i++) {
    if (rows[i].name === MARKER) { markerIdx = i; break; }
  }
  const endIdx = markerIdx >= 0 ? markerIdx : rows.length;
  let docIdx = -1;
  for (let i=0;i<endIdx;i++) {
    if (rows[i].name.toLowerCase().includes('документация') && !rows[i].des) { docIdx = i; break; }
  }
  if (docIdx < 0) throw new Error('Раздел «Документация» не найден');
  let rootDes = null;
  for (let i=docIdx+1;i<endIdx;i++) {
    const r = rows[i];
    if (!r.name && !r.des) continue;
    if (detectSection(r.name) && !r.des) break;
    if (r.des && r.des.trim() !== '') { rootDes = r.des.trim(); break; }
  }
  if (!rootDes) throw new Error('Корневое обозначение не найдено');
  rootDes = baseDesignation(rootDes);
  const constantItems = collectItems(rows, docIdx+1, endIdx);
  if (markerIdx < 0) return [ makeTree(rootDes, constantItems) ];
  const blocks = [];
  let current = null, currentSection = null, lastFullDes = '', lastFullDesBorrowed = false;
  for (let i=markerIdx+1;i<rows.length;i++) {
    const r = rows[i];
    if (!r.name && !r.des) continue;
    const sec = detectSection(r.name);
    let headerDes = null, refTo = null;
    if (!r.des && r.name && !sec) {
      const s = r.name.trim();
      const lower = s.toLowerCase();
      const hasToZhe = lower.includes('то же');
      if (hasToZhe) {
        const matches = s.match(DES_ANYWHERE_RE) || [];
        if (matches.length >= 1) headerDes = matches[0].trim();
        if (matches.length >= 2) refTo = matches[matches.length - 1].trim();
      } else if (BLOCK_HEADER_RE.test(s)) {
        headerDes = s;
      }
    }
    if (headerDes) {
      current = { header: headerDes, refTo: refTo, items: [] };
      blocks.push(current);
      currentSection = null; lastFullDes = ''; lastFullDesBorrowed = false;
      continue;
    }
    if (sec && !r.des) { currentSection = sec; lastFullDes = ''; lastFullDesBorrowed = false; continue; }
    if (!current || !currentSection) continue;
    if (IGNORE_SECTIONS.includes(currentSection)) continue;
    if (r.des && r.des.trim() !== '') {
      let des = r.des.trim();
      const count = parseFloat(r.count) || 1;
      const ownBorrowed = !!(r.note && r.note.includes('▼'));
      let borrowed;
      if (SUFFIX_ONLY_RE.test(des)) {
        if (lastFullDes) {
          const base = lastFullDes.replace(/-\d+$/, '');
          des = base + des;
        }
        borrowed = lastFullDesBorrowed;
      } else {
        lastFullDes = des;
        lastFullDesBorrowed = ownBorrowed;
        borrowed = ownBorrowed;
      }
      current.items.push({ des, name: r.name, count, borrowed, section: currentSection });
    } else if (r.name && r.name.trim() !== '') {
      if (current.items.length > 0) current.items[current.items.length-1].name += ' ' + r.name.trim();
    }
  }
  resolveReferences(blocks);
  if (blocks.length === 0) return [ makeTree(rootDes, constantItems) ];
  const result = [];
  for (const b of blocks) {
    const combined = constantItems.concat(b.items);
    result.push( makeTree(b.header, combined) );
  }
  return result;
}

function resolveReferences(blocks) {
  const map = new Map();
  for (const b of blocks) map.set(b.header, b);
  for (let pass = 0; pass < 10; pass++) {
    let changed = false;
    for (const b of blocks) {
      if (!b.refTo) continue;
      if (b.items.length > 0) continue;
      const src = map.get(b.refTo);
      if (src && src.items.length > 0) {
        b.items = src.items.map(it => ({
          des: it.des, name: it.name, count: it.count,
          borrowed: it.borrowed, section: it.section
        }));
        changed = true;
      }
    }
    if (!changed) break;
  }
}

function collectItems(rows, from, to) {
  const items = [];
  let currentSection = null, lastFullDes = '', lastFullDesBorrowed = false;
  for (let i=from;i<to;i++) {
    const r = rows[i];
    if (!r.name && !r.des) continue;
    const sec = detectSection(r.name);
    if (sec && !r.des) { currentSection = sec; lastFullDes = ''; lastFullDesBorrowed = false; continue; }
    if (!currentSection) continue;
    if (IGNORE_SECTIONS.includes(currentSection)) continue;
    if (r.des && r.des.trim() !== '') {
      let des = r.des.trim();
      const count = parseFloat(r.count) || 1;
      const ownBorrowed = !!(r.note && r.note.includes('▼'));
      let borrowed;
      if (SUFFIX_ONLY_RE.test(des)) {
        if (lastFullDes) {
          const base = lastFullDes.replace(/-\d+$/, '');
          des = base + des;
        }
        borrowed = lastFullDesBorrowed;
      } else {
        lastFullDes = des;
        lastFullDesBorrowed = ownBorrowed;
        borrowed = ownBorrowed;
      }
      items.push({ des, name: r.name, count, borrowed, section: currentSection });
    } else if (r.name && r.name.trim() !== '') {
      if (items.length > 0) items[items.length-1].name += ' ' + r.name.trim();
    }
  }
  return items;
}

function makeTree(rootDes, items) {
  const root = {
    designation: rootDes, name: '', count: 1, totalCount: 1,
    borrowed: false, isNew: false, isManual: false,
    type: 'Сборка', section: null,
    isExpanded: true, hidden: false, children: []
  };
  for (const it of items) {
    let type = 'Деталь';
    if (it.section === 'Сборочные единицы') type = 'Сборка';
    else if (it.section === 'Жгуты') type = 'Жгут';
    root.children.push({
      designation: it.des, name: it.name,
      count: it.count, totalCount: it.count,
      note: it.borrowed ? '' : 'НОВОЕ!',
      borrowed: it.borrowed, isNew: !it.borrowed, isManual: false,
      type, section: it.section,
      isExpanded: false, hidden: false, children: []
    });
  }
  return root;
}

async function applySubFiles() {
  if (!subFiles.length) return;
  const map = new Map();
  for (const sf of subFiles) {
    try {
      const rows = await parseAnyFile(sf.buffer, sf.name);
      const subTrees = buildTreesFromRows(rows);
      for (const t of subTrees) {
        const exactKey = String(t.designation || '').trim();
        if (exactKey) map.set(exactKey, t);
        const baseKey = baseDesignation(t.designation);
        if (baseKey && baseKey !== exactKey) map.set(baseKey, t);
      }
    } catch (e) {}
  }
  function expandNode(node) {
    if (!node.children || !node.children.length) return;
    for (const child of node.children) {
      if (child.type !== 'Сборка') continue;
      const exactKey = String(child.designation || '').trim();
      const baseKey = baseDesignation(child.designation);
      const sub = map.get(exactKey) || map.get(baseKey);
      if (sub) {
        child.children = sub.children.map(cloneNode);
        child.isExpanded = false;
        expandNode(child);
      }
    }
  }
  for (const tree of trees) expandNode(tree);
}

function cloneNode(n) {
  return {
    designation: n.designation, name: n.name,
    count: n.count, totalCount: n.count,
    note: n.note || '',
    borrowed: n.borrowed, isNew: !n.borrowed, isManual: false,
    type: n.type, section: n.section,
    isExpanded: false, hidden: false,
    children: (n.children||[]).map(cloneNode)
  };
}

function postProcess() {
  for (const tree of trees) recalcTotals(tree, 1);
}
function recalcTotals(node, parentTotal) {
  node.totalCount = (node.count || 1) * parentTotal;
  for (const c of node.children) recalcTotals(c, node.totalCount);
}

// ПРЕДПРОСМОТР
function rebuildPreview() {
  flatNodes = [];
  if (!trees.length) { $('previewList').innerHTML = ''; return; }
  let html = '<div class="preview-header-row">' +
    '<span class="h-del"></span>' +
    '<span class="h-cb"></span><span class="h-dot"></span>' +
    '<span class="h-num">№</span><span class="h-name">Наименование</span>' +
    '<span class="h-des">Обозначение</span>' +
    '<span class="h-local"></span><span class="h-total">Кол.</span>' +
    '<span class="h-note">Примечание</span></div>';
  trees.forEach((tree, ti) => {
    html += '<div class="tree-separator">🌳 ' + escapeHtml(tree.designation) + '</div>';
    flatten(tree, ti, '', 0, flatNodes, true);
  });
  for (let i = 0; i < flatNodes.length; i++) {
    const info = flatNodes[i];
    if (info.isAddRow) html += renderAddRow(info.treeIdx, info.parentNode, info.depth);
    else html += renderRow(info, i);
  }
  $('previewList').innerHTML = html;
  attachHandlers();
}

function flatten(node, treeIdx, prefix, depth, arr, isRoot) {
  arr.push({ node, treeIdx, prefix, depth, isRoot });
  if (node.isExpanded && node.children && node.children.length) {
    let c = 0;
    for (const child of node.children) {
      c++;
      flatten(child, treeIdx, prefix + c + '.', depth+1, arr, false);
    }
    arr.push({ isAddRow: true, treeIdx, parentNode: node, depth: depth });
  }
}

function renderRow(info, idx) {
  const { node, prefix, depth } = info;
  const num = prefix || '1';
  const indent = '&nbsp;'.repeat(depth * 3);
  const canExpand = node.children && node.children.length > 0;
  const dotClass = canExpand ? (node.isExpanded ? 'filled' : 'empty') : 'disabled';
  const checkedAttr = node.hidden ? '' : 'checked';
  const delBtn = node.isManual
    ? '<button data-action="delete" data-idx="' + idx + '" title="Удалить">✕</button>' : '';
  return '<div class="preview-item" data-idx="' + idx + '">' +
    '<span class="item-del">' + delBtn + '</span>' +
    '<input type="checkbox" ' + checkedAttr + ' class="vis-cb">' +
    '<span class="expand-dot ' + dotClass + '" data-idx="' + idx + '"></span>' +
    '<span class="item-number">' + indent + escapeHtml(num) + '</span>' +
    '<span class="item-name"><input class="edit-field" data-field="name" value="' + escapeHtml(node.name) + '"></span>' +
    '<span class="item-des"><input class="edit-field" data-field="des" value="' + escapeHtml(node.designation) + '"></span>' +
    '<span class="item-local"><input class="edit-field" type="number" step="any" data-field="count" value="' + (node.count||1) + '"></span>' +
    '<span class="item-total">' + (node.totalCount||0) + '</span>' +
    '<span class="item-note"><input class="edit-field" data-field="note" value="' + escapeHtml(node.note || '') + '"></span>' +
    '</div>';
}

function renderAddRow(treeIdx, parentNode, depth) {
  const parentId = parentNode ? flatNodes.findIndex(f => f.node === parentNode) : -1;
  return '<div class="add-row-line">' +
    '<span class="item-del"></span>' +
    '<button class="add-btn" data-action="add" data-tree="' + treeIdx + '" data-parent="' + parentId + '">➕ Добавить строку</button>' +
    '</div>';
}

function attachHandlers() {
  const list = $('previewList');
  list.querySelectorAll('.vis-cb').forEach(cb => {
    cb.addEventListener('change', function(){
      const idx = parseInt(this.closest('.preview-item').dataset.idx);
      flatNodes[idx].node.hidden = !this.checked;
    });
  });
  list.querySelectorAll('.expand-dot').forEach(dot => {
    dot.addEventListener('click', function(){
      const idx = parseInt(this.dataset.idx);
      const node = flatNodes[idx].node;
      if (node && node.children && node.children.length) {
        node.isExpanded = !node.isExpanded;
        rebuildPreview();
      }
    });
  });
  list.querySelectorAll('.edit-field').forEach(inp => {
    inp.addEventListener('input', function(){
      const idx = parseInt(this.closest('.preview-item').dataset.idx);
      const node = flatNodes[idx].node;
      const f = this.dataset.field;
      if (f === 'name') node.name = this.value;
      else if (f === 'des') node.designation = this.value;
      else if (f === 'count') {
        node.count = parseFloat(this.value) || 1;
        postProcess();
        refreshTotals();
      }
      else if (f === 'note') node.note = this.value;
    });
  });
  list.querySelectorAll('button[data-action="delete"]').forEach(btn => {
    btn.addEventListener('click', function(){
      const idx = parseInt(this.dataset.idx);
      const info = flatNodes[idx];
      if (!info || !info.node || !info.node.isManual) return;
      removeNodeFromTree(info.treeIdx, info.node);
      postProcess();
      rebuildPreview();
    });
  });
  list.querySelectorAll('button[data-action="add"]').forEach(btn => {
    btn.addEventListener('click', function(){
      const treeIdx = parseInt(this.dataset.tree);
      const parentId = parseInt(this.dataset.parent);
      let parentNode = null;
      if (parentId >= 0 && flatNodes[parentId] && flatNodes[parentId].node) {
        parentNode = flatNodes[parentId].node;
      }
      const newNode = {
        designation: '', name: '', count: 1, totalCount: 1,
        note: '', borrowed: false, isNew: true, isManual: true,
        type: 'Деталь', section: null,
        isExpanded: false, hidden: false, children: []
      };
      if (parentNode) {
        parentNode.children.push(newNode);
        parentNode.isExpanded = true;
      } else {
        trees[treeIdx].children.push(newNode);
      }
      postProcess();
      rebuildPreview();
    });
  });
}

function removeNodeFromTree(treeIdx, node) {
  const tree = trees[treeIdx];
  if (!tree || tree === node) return;
  function tryRemove(parent) {
    if (!parent.children) return false;
    const idx = parent.children.indexOf(node);
    if (idx >= 0) { parent.children.splice(idx, 1); return true; }
    for (const c of parent.children) if (tryRemove(c)) return true;
    return false;
  }
  tryRemove(tree);
}

function refreshTotals() {
  $('previewList').querySelectorAll('.preview-item').forEach(el => {
    const idx = parseInt(el.dataset.idx);
    const node = flatNodes[idx].node;
    const cell = el.querySelector('.item-total');
    if (cell) cell.textContent = node.totalCount || 0;
  });
}

function rebuildOrderTreeSelect() {
  const sel = $('orderTreeSelect');
  sel.innerHTML = '';
  if (!trees.length) {
    $('orderTreePicker').style.display = 'none';
    $('generateOrderBtn').disabled = true;
    return;
  }
  trees.forEach((t, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = t.designation || ('Дерево ' + (i+1));
    sel.appendChild(opt);
  });
  sel.value = '0';
  $('orderTreePicker').style.display = 'flex';
  $('generateOrderBtn').disabled = false;
}

// ОТЧЁТ
function generateReport() {
  const wb = XLSX.utils.book_new();
  const usedNames = {};
  for (const tree of trees) {
    const rows = [['№','Наименование','Обозначение','Кол.','Примечание']];
    let counter = 0;
    function addNode(node, prefix, isRoot) {
      if (node.hidden) return;
      counter++;
      const num = prefix + counter;
      rows.push([num, node.name||'', node.designation || '', node.totalCount || 0, isRoot ? '' : (node.note || '')]);
      if (node.isExpanded && node.children && node.children.length) {
        const saved = counter;
        counter = 0;
        for (const c of node.children) addNode(c, num + '.', false);
        counter = saved;
      }
    }
    addNode(tree, '', true);
    let sheetName = (tree.designation || 'Лист').replace(/[\\/*?:\[\]]/g,'_').slice(0,31);
    if (!sheetName) sheetName = 'Лист';
    let base = sheetName, k = 1;
    while (usedNames[sheetName]) { sheetName = (base.slice(0,28) + '_' + (++k)).slice(0,31); }
    usedNames[sheetName] = true;
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:14},{wch:45},{wch:28},{wch:8},{wch:24}];
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }
  return XLSX.write(wb, { bookType:'xlsx', type:'array' });
}

// ЗАЯВКА
async function generateOrder(configName, srokValue) {
  const idx = parseInt($('orderTreeSelect').value);
  if (isNaN(idx) || idx < 0 || idx >= trees.length) { alert('Выберите дерево'); return; }
  const tree = trees[idx];
  const cfg = configName && configs[configName] ? configs[configName] : emptyConfig();

  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Заявка');
    ws.columns = [
      { width: 6 }, { width: 30 }, { width: 10 }, { width: 10 }, { width: 10 },
      { width: 20 }, { width: 8 }, { width: 25 }, { width: 15 }
    ];
    const thinBorder = {
      top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'}
    };

    ws.mergeCells('A1:I1');
    ws.getCell('A1').value = 'Заявка';
    ws.getCell('A1').font = { bold: true, size: 14 };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    ws.mergeCells('A3:I3');
    ws.getCell('A3').value = 'Изделие: ' + (cfg.izdelie || '');

    ws.mergeCells('A4:F4');
    ws.getCell('A4').value = 'Децимальный номер изделия: ' + (cfg.decimal || '');
    ws.mergeCells('G4:I4');
    ws.getCell('G4').value = 'Дата: ' + (cfg.data || '');

    ws.mergeCells('A5:F5');
    ws.getCell('A5').value = 'Договор: ' + (cfg.dogovor || '');
    ws.mergeCells('G5:I5');
    ws.getCell('G5').value = cfg.ooo || '';

    ws.mergeCells('A6:I6');
    ws.getCell('A6').value = 'Приемка: ' + (cfg.priemka || '');
    ws.mergeCells('A7:I7');
    ws.getCell('A7').value = 'Гарантия: ' + (cfg.garantiya || '');
    ws.mergeCells('A8:I8');
    ws.getCell('A8').value = 'Поставщик: ' + (cfg.postavshik || '');
    ws.mergeCells('A9:I9');
    ws.getCell('A9').value = 'Заказчик: ' + (cfg.zakazchik || '');

    ws.mergeCells('A10:D10');
    ws.getCell('A10').value = 'Ведущий: ' + (cfg.vedushiy || '');
    ws.mergeCells('E10:I10');
    ws.getCell('E10').value = 'конструктор: ' + (cfg.konstruktor || '');

    ws.getRow(11).getCell(1).value = '№';
    ws.getRow(11).getCell(2).value = 'Наименование/обозначение';
    ws.getRow(11).getCell(6).value = 'Дец. номер/артикул';
    ws.getRow(11).getCell(7).value = 'Кол-во';
    ws.getRow(11).getCell(8).value = 'Примечание';
    ws.getRow(11).getCell(9).value = 'Срок';
    ws.mergeCells('B11:E11');
    for (let c=1;c<=9;c++) {
      const cell = ws.getRow(11).getCell(c);
      cell.font = { bold: true };
      cell.border = thinBorder;
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
    }

    const dataRows = [];
    let counter = 0;
    function collect(node, prefix, isRoot) {
      if (node.hidden) return;
      counter++;
      const num = prefix + counter;
      const note = isRoot ? '' : (node.borrowed ? '' : 'НОВОЕ!');
      dataRows.push({
        num: num, name: node.name || '',
        des: node.designation || '',
        count: node.totalCount || 0, note: note
      });
      if (node.isExpanded && node.children && node.children.length) {
        const saved = counter;
        counter = 0;
        for (const c of node.children) collect(c, num + '.', false);
        counter = saved;
      }
    }
    collect(tree, '', true);

    const START_ROW = 12;
    dataRows.forEach((row, i) => {
      const rowIdx = START_ROW + i;

      // Объединяем B:E в этой строке
      ws.mergeCells('B' + rowIdx + ':E' + rowIdx);

      const r = ws.getRow(rowIdx);
      r.getCell(1).value = row.num;
      r.getCell(2).value = row.name;
      r.getCell(2).alignment = { horizontal:'left', vertical:'middle', wrapText: true };
      r.getCell(6).value = row.des;
      r.getCell(7).value = row.count;
      if (row.note) {
        r.getCell(8).value = row.note;
        r.getCell(8).alignment = { horizontal:'center', vertical:'middle' };
      }
      for (let c=1;c<=9;c++) r.getCell(c).border = thinBorder;
      r.commit();
    });

    const END_ROW = START_ROW + dataRows.length - 1;
    if (dataRows.length > 0) {
      if (END_ROW > START_ROW) ws.mergeCells('I' + START_ROW + ':I' + END_ROW);
      ws.getCell('I' + START_ROW).value = srokValue || '';
      ws.getCell('I' + START_ROW).alignment = { horizontal:'center', vertical:'middle' };
    }

    const SIGN_START = END_ROW + 2;
    ws.mergeCells('A' + SIGN_START + ':I' + SIGN_START);
    ws.getCell('A' + SIGN_START).value = 'Инженер-конструктор: ' + (cfg.inzhKonstr || '') + '  /______________/';
    ws.mergeCells('A' + (SIGN_START+1) + ':I' + (SIGN_START+1));
    ws.getCell('A' + (SIGN_START+1)).value = 'Ведущий инженер: ' + (cfg.vedInzh || '') + '  /______________/';
    ws.mergeCells('A' + (SIGN_START+2) + ':I' + (SIGN_START+2));
    ws.getCell('A' + (SIGN_START+2)).value = 'Начальник отдела: ' + (cfg.nachOtdela || '') + '  /______________/';

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type:'application/octet-stream' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (tree.designation || 'заявка').replace(/[\\/*?:"<>|]/g,'_') + '_заявка.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    log('✅ Заявка сформирована: ' + tree.designation, 'ok');
  } catch (e) {
    log('❌ Ошибка заявки: ' + e.message, 'err');
    console.error(e);
  }
}

// ПРЕДПРОСМОТР — поля конфига на сессию
function rebuildPreviewConfigFields() {
  const grid = $('previewConfigGrid');
  grid.innerHTML = '';
  const base = previewSelectedConfigName && configs[previewSelectedConfigName]
    ? configs[previewSelectedConfigName]
    : emptyConfig();
  orderSessionConfig = Object.assign({}, base);
  CONFIG_FIELDS.forEach(f => {
    const div = document.createElement('div');
    div.className = 'field';
    div.innerHTML = '<label>' + escapeHtml(f.label) + '</label>' +
                    '<input type="text" data-key="' + f.key + '" value="' + escapeHtml(orderSessionConfig[f.key]||'') + '">';
    grid.appendChild(div);
  });
  grid.querySelectorAll('input[data-key]').forEach(inp => {
    inp.addEventListener('input', function(){
      orderSessionConfig[this.dataset.key] = this.value;
    });
  });
}

// ЗАГРУЗКА ФАЙЛОВ
async function handleMain(file) {
  if (!file) return;
  const lower = file.name.toLowerCase();
  if (!lower.endsWith('.xlsx') && !lower.endsWith('.docx')) { alert('Только .xlsx или .docx'); return; }
  clearLog();
  $('buildBtn').disabled = true;
  try {
    mainBuffer = await file.arrayBuffer();
    mainName = file.name;
    $('mainFileName').textContent = file.name;
    $('mainFileSize').textContent = '(' + (file.size/1024).toFixed(1) + ' KB)';
    $('mainFileInfo').classList.add('show');
    const rows = await parseAnyFile(mainBuffer, file.name);
    trees = buildTreesFromRows(rows);
    log('✅ Деревьев: ' + trees.length, 'ok');
    trees.forEach((t,i) => log('   [' + (i+1) + '] ' + t.designation, 'inf'));
    subFiles = []; renderSubFiles();
    $('buildBtn').disabled = false;
  } catch (e) {
    log('❌ ' + e.message, 'err');
    $('buildBtn').disabled = true;
  }
}

async function handleSubFiles(files) {
  if (!files.length) return;
  for (const f of files) {
    const lower = f.name.toLowerCase();
    if (!lower.endsWith('.xlsx') && !lower.endsWith('.docx')) continue;
    try {
      const buffer = await f.arrayBuffer();
      subFiles.push({ name: f.name, buffer });
      log('📎 Подсборка: ' + f.name, 'ok');
    } catch (e) {}
  }
  renderSubFiles();
}

function renderSubFiles() {
  const box = $('subFilesList');
  if (!subFiles.length) { box.innerHTML = '<span style="color:var(--text-mute);">Нет подсборок</span>'; return; }
  box.innerHTML = subFiles.map((sf,i) =>
    '<div class="sub-file-item"><span>📎 ' + escapeHtml(sf.name) + '</span>' +
    '<button class="remove-sub" data-index="' + i + '">✕</button></div>'
  ).join('');
  box.querySelectorAll('.remove-sub').forEach(b => {
    b.onclick = () => { subFiles.splice(parseInt(b.dataset.index),1); renderSubFiles(); };
  });
}

function setupDropZone(dz, inp, multi, cb) {
  dz.addEventListener('click', () => inp.click());
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('dragover');
    const fs = e.dataTransfer.files;
    if (fs.length) cb(multi ? Array.from(fs) : fs[0]);
  });
  inp.addEventListener('change', () => {
    if (inp.files.length) cb(multi ? Array.from(inp.files) : inp.files[0]);
  });
}
setupDropZone($('mainDropZone'), $('mainFileInput'), false, handleMain);
setupDropZone($('subDropZone'), $('subFileInput'), true, handleSubFiles);

$('removeMainFile').addEventListener('click', () => {
  mainBuffer = null; mainName = '';
  $('mainFileInfo').classList.remove('show');
  $('mainFileInput').value = '';
  trees = []; subFiles = []; renderSubFiles();
  $('buildBtn').disabled = true;
});

// КНОПКИ
$('buildBtn').addEventListener('click', async () => {
  if (!trees.length) { alert('Загрузите главный файл'); return; }
  $('buildBtn').disabled = true;
  try {
    const rows = await parseAnyFile(mainBuffer, mainName);
    trees = buildTreesFromRows(rows);
    await applySubFiles();
    postProcess();
    rebuildPreview();
    rebuildOrderTreeSelect();
    previewSelectedConfigName = '';
    $('previewConfigSelect').value = '';
    rebuildPreviewConfigFields();
    $('previewSrokInput').value = '';
    $('previewModal').classList.add('show');
  } catch (e) {
    log('❌ ' + e.message, 'err');
  } finally {
    $('buildBtn').disabled = false;
  }
});

$('closePreview').addEventListener('click', () => $('previewModal').classList.remove('show'));
$('previewConfigHeader').addEventListener('click', () => {
  $('previewConfig').classList.toggle('open');
});

$('previewConfigSelect').addEventListener('change', function() {
  previewSelectedConfigName = this.value;
  rebuildPreviewConfigFields();
});

$('downloadReportBtn').addEventListener('click', () => {
  try {
    const arr = generateReport();
    const blob = new Blob([arr], { type:'application/octet-stream' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (trees[0]?.designation || 'отчёт').replace(/[\\/*?:"<>|]/g,'_') + '_отчет.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    log('✅ Отчёт скачан', 'ok');
  } catch (e) { log('❌ ' + e.message, 'err'); }
});

$('generateOrderBtn').addEventListener('click', () => {
  // Если конфиг выбран в предпросмотре — генерируем сразу
  if (previewSelectedConfigName) {
    generateOrder(previewSelectedConfigName, $('previewSrokInput').value || '');
    return;
  }
  // Иначе открываем окно выбора
  rebuildOrderConfigSelect();
  $('orderSrokInput').value = $('previewSrokInput').value || '';
  $('orderConfigModal').classList.add('show');
});

$('closeOrderConfigModal').addEventListener('click', () => $('orderConfigModal').classList.remove('show'));

$('confirmGenerateOrderBtn').addEventListener('click', async () => {
  const cfgName = $('orderConfigSelect').value;
  const srok = $('orderSrokInput').value;
  $('orderConfigModal').classList.remove('show');
  await generateOrder(cfgName, srok);
});

$('expandAllBtn').addEventListener('click', () => {
  function exp(n){ n.isExpanded = true; n.children.forEach(exp); }
  trees.forEach(exp);
  rebuildPreview();
});
$('collapseAllBtn').addEventListener('click', () => {
  function col(n){ n.isExpanded = false; n.children.forEach(col); }
  trees.forEach(col);
  rebuildPreview();
});

loadConfigsFromLS();
refreshConfigUI();
rebuildOrderConfigSelect();

})();
</script>
</body>
</html>
