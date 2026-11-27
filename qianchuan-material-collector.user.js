// ==UserScript==
// @name         千川素材请求采集
// @namespace    VideoCollection
// @version      1.6.6
// @description  采集千川素材与抖店自营视频，支持单页 CSV 与双页 Excel 汇总导出。
// @updateURL    https://raw.githubusercontent.com/xiaomingxinggit/VideoCollection/main/qianchuan-material-collector.user.js
// @downloadURL  https://raw.githubusercontent.com/xiaomingxinggit/VideoCollection/main/qianchuan-material-collector.user.js
//
// @match        https://qianchuan.jinritemai.com/dataV3/roi2-material-analysis*
// @match        https://compass.jinritemai.com/shop/video/self*
// @require      https://cdn.jsdelivr.net/npm/vue@3.5.46/dist/vue.global.prod.js#sha256-wuS4iOy7EseyT9Xov2VIPOqrc8N7eUAveWgCOXftPoE=
// @require      data:text/javascript;base64,Y29uc3QgX192Y0dsb2JhbCA9IGdsb2JhbFRoaXM7IGNvbnN0IF9fdmNQYWdlR2xvYmFsID0gdHlwZW9mIHVuc2FmZVdpbmRvdyA9PT0gJ29iamVjdCcgPyB1bnNhZmVXaW5kb3cgOiBfX3ZjR2xvYmFsOyBjb25zdCBfX3ZjSGFkVnVlID0gT2JqZWN0Lmhhc093bihfX3ZjR2xvYmFsLCAnVnVlJyk7IGNvbnN0IF9fdmNPbGRWdWUgPSBfX3ZjR2xvYmFsLlZ1ZTsgY29uc3QgX192Y0hhZEVsZW1lbnRQbHVzID0gT2JqZWN0Lmhhc093bihfX3ZjR2xvYmFsLCAnRWxlbWVudFBsdXMnKTsgY29uc3QgX192Y09sZEVsZW1lbnRQbHVzID0gX192Y0dsb2JhbC5FbGVtZW50UGx1czsgY29uc3QgX192Y0hhZFBhZ2VWdWUgPSBPYmplY3QuaGFzT3duKF9fdmNQYWdlR2xvYmFsLCAnVnVlJyk7IGNvbnN0IF9fdmNPbGRQYWdlVnVlID0gX192Y1BhZ2VHbG9iYWwuVnVlOyBjb25zdCBfX3ZjSGFkUGFnZUVsZW1lbnRQbHVzID0gT2JqZWN0Lmhhc093bihfX3ZjUGFnZUdsb2JhbCwgJ0VsZW1lbnRQbHVzJyk7IGNvbnN0IF9fdmNPbGRQYWdlRWxlbWVudFBsdXMgPSBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1czsgX192Y0dsb2JhbC5WdWUgPSBWdWU7IF9fdmNQYWdlR2xvYmFsLlZ1ZSA9IFZ1ZTs=
// @require      https://cdn.jsdelivr.net/npm/element-plus@2.14.8/dist/index.full.min.js#sha256-D/7wpbpse4CCbLyQwzY2KGIjP5bvtkEHnzgXXJyOebI=
// @require      data:text/javascript;base64,Y29uc3QgRWxlbWVudFBsdXMgPSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzOyBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1cyA9IEVsZW1lbnRQbHVzOyBjb25zdCBfX3ZjUmVzdG9yZUZyYW1ld29ya0dsb2JhbHMgPSAoKSA9PiB7IGlmIChfX3ZjSGFkVnVlKSBfX3ZjR2xvYmFsLlZ1ZSA9IF9fdmNPbGRWdWU7IGVsc2UgZGVsZXRlIF9fdmNHbG9iYWwuVnVlOyBpZiAoX192Y0hhZEVsZW1lbnRQbHVzKSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzID0gX192Y09sZEVsZW1lbnRQbHVzOyBlbHNlIGRlbGV0ZSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzOyBpZiAoX192Y0hhZFBhZ2VWdWUpIF9fdmNQYWdlR2xvYmFsLlZ1ZSA9IF9fdmNPbGRQYWdlVnVlOyBlbHNlIGRlbGV0ZSBfX3ZjUGFnZUdsb2JhbC5WdWU7IGlmIChfX3ZjSGFkUGFnZUVsZW1lbnRQbHVzKSBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1cyA9IF9fdmNPbGRQYWdlRWxlbWVudFBsdXM7IGVsc2UgZGVsZXRlIF9fdmNQYWdlR2xvYmFsLkVsZW1lbnRQbHVzOyB9Ow==
// @require      data:text/javascript;base64,Y29uc3QgX192Y1NoZWV0UGFnZUdsb2JhbCA9IHR5cGVvZiB1bnNhZmVXaW5kb3cgPT09ICdvYmplY3QnID8gdW5zYWZlV2luZG93IDogbnVsbDsgY29uc3QgX192Y0hhZFBhZ2VYbHN4ID0gISFfX3ZjU2hlZXRQYWdlR2xvYmFsICYmIE9iamVjdC5oYXNPd24oX192Y1NoZWV0UGFnZUdsb2JhbCwgJ1hMU1gnKTsgY29uc3QgX192Y09sZFBhZ2VYbHN4ID0gX192Y1NoZWV0UGFnZUdsb2JhbD8uWExTWDs=
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.mini.min.js#sha256-DLNT+DDXKIOFSSyD0newWN3qxmTKUc8Tk6of0+K3CTk=
// @require      data:text/javascript;base64,aWYgKF9fdmNTaGVldFBhZ2VHbG9iYWwpIHsgaWYgKF9fdmNIYWRQYWdlWGxzeCkgX192Y1NoZWV0UGFnZUdsb2JhbC5YTFNYID0gX192Y09sZFBhZ2VYbHN4OyBlbHNlIGRlbGV0ZSBfX3ZjU2hlZXRQYWdlR2xvYmFsLlhMU1g7IH0=
// @resource     ELEMENT_PLUS_CSS https://cdn.jsdelivr.net/npm/element-plus@2.14.8/dist/index.css#sha256-9vVHqUarXedNRDZ6BVb0lN1jDGdZmtT2amrCAKhegKY=
// @resource     ELEMENT_PLUS_DARK_CSS https://cdn.jsdelivr.net/npm/element-plus@2.14.8/theme-chalk/dark/css-vars.css#sha256-KOclPpHatx0lmXDXwrIWdM7/op+kRvWWikCTMhUWFgA=
// @run-at       document-start
// @sandbox      raw
// @noframes
// @grant        GM_registerMenuCommand
// @grant        GM_getResourceText
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        unsafeWindow
// ==/UserScript==

const videoCollectionWorkbook = (() => {
  'use strict';

  const schemaVersion = 1;
  const storageKeys = Object.freeze({
    qianchuan: 'video-collection:qianchuan:v1',
    compass: 'video-collection:compass:v1',
  });
  const labels = Object.freeze({ qianchuan: '千川素材', compass: '自营视频' });
  const countFormat = new Intl.NumberFormat('zh-CN');
  const maxPairAgeMs = 24 * 60 * 60 * 1000;

  function isDate(value) {
    const text = String(value);
    return /^\d{4}-\d{2}-\d{2}$/.test(text)
      && new Date(`${text}T00:00:00Z`).toJSON()?.slice(0, 10) === text;
  }

  function validRow(source, row) {
    if (!row || typeof row !== 'object') return false;
    if (source === 'qianchuan') {
      return typeof row.promotionType === 'string' && typeof row.materialId === 'string'
        && typeof row.title === 'string' && Number.isFinite(row.spendYuan) && row.spendYuan >= 0
        && Number.isSafeInteger(row.impressions) && row.impressions >= 0;
    }
    return source === 'compass' && typeof row.title === 'string'
      && Number.isSafeInteger(row.watchCount) && row.watchCount >= 0;
  }

  function canonicalRows(source, rows) {
    if (!Array.isArray(rows)) return null;
    const safeRows = rows.map((row) => source === 'qianchuan' ? {
      promotionType: row?.promotionType,
      materialId: row?.materialId,
      title: row?.title,
      spendYuan: row?.spendYuan,
      impressions: row?.impressions,
    } : {
      title: row?.title,
      watchCount: row?.watchCount,
    });
    return safeRows.every((row) => validRow(source, row)) ? safeRows : null;
  }

  function scopeFingerprint(value) {
    let hash = 2166136261;
    for (const character of String(value)) {
      hash ^= character.codePointAt(0);
      hash = Math.imul(hash, 16777619);
    }
    return `v1:${(hash >>> 0).toString(36)}`;
  }

  function normalizeSnapshot(source, value) {
    if (!value || value.schemaVersion !== schemaVersion || value.source !== source
      || value.complete !== true || !isDate(value.startDate) || !isDate(value.endDate)
      || value.startDate > value.endDate || !/^v1:[a-z0-9]+$/.test(value.scope)
      || !Number.isFinite(Date.parse(value.collectedAt))
      || !Number.isSafeInteger(value.scanned) || value.scanned < 0
      || !Array.isArray(value.rows) || !value.rows.every((row) => validRow(source, row))) {
      return null;
    }
    return value;
  }

  function read(source) {
    const key = storageKeys[source];
    if (!key || typeof GM_getValue !== 'function') return null;
    try {
      const stored = GM_getValue(key, '');
      const value = typeof stored === 'string' ? JSON.parse(stored || 'null') : stored;
      return normalizeSnapshot(source, value);
    } catch {
      return null;
    }
  }

  function invalidate(source) {
    const key = storageKeys[source];
    if (!key || typeof GM_deleteValue !== 'function') return false;
    try {
      GM_deleteValue(key);
      return true;
    } catch {
      return false;
    }
  }

  function begin(source, scope, period) {
    if (!storageKeys[source] || typeof scope !== 'string' || !scope
      || !isDate(period?.startDate) || !isDate(period?.endDate)) return false;
    const otherSource = source === 'qianchuan' ? 'compass' : 'qianchuan';
    const current = read(source);
    const other = read(otherSource);
    const scopeChanged = current && current.scope !== scopeFingerprint(scope);
    const periodChanged = other && (other.startDate !== period.startDate || other.endDate !== period.endDate);
    if (scopeChanged || periodChanged || summary().ready) invalidate(otherSource);
    invalidate(source);
    return true;
  }

  function save(source, snapshot) {
    const key = storageKeys[source];
    const rows = canonicalRows(source, snapshot?.rows);
    if (!key || typeof GM_setValue !== 'function' || !snapshot
      || typeof snapshot.scope !== 'string' || !snapshot.scope || !rows) {
      invalidate(source);
      return false;
    }
    try {
      const value = {
        schemaVersion,
        source,
        complete: true,
        scope: scopeFingerprint(snapshot.scope),
        startDate: snapshot.startDate,
        endDate: snapshot.endDate,
        collectedAt: new Date().toISOString(),
        scanned: Number.isSafeInteger(snapshot.scanned) && snapshot.scanned >= 0
          ? snapshot.scanned : rows.length,
        rows,
      };
      if (!normalizeSnapshot(source, value)) {
        invalidate(source);
        return false;
      }
      GM_setValue(key, JSON.stringify(value));
      return true;
    } catch {
      invalidate(source);
      return false;
    }
  }

  function periodText(snapshot) {
    return `${snapshot.startDate} 至 ${snapshot.endDate}`;
  }

  function statusItem(source, snapshot) {
    return snapshot ? {
      key: source,
      label: labels[source],
      state: 'done',
      text: `已采集 ${countFormat.format(snapshot.rows.length)} 条 ${periodText(snapshot)}`,
    } : {
      key: source,
      label: labels[source],
      state: 'waiting',
      text: '未采集',
    };
  }

  function summary() {
    const qianchuan = read('qianchuan');
    const compass = read('compass');
    const complete = Boolean(qianchuan && compass);
    const samePeriod = complete && qianchuan.startDate === compass.startDate
      && qianchuan.endDate === compass.endDate;
    const closeInTime = complete && Math.abs(
      Date.parse(qianchuan.collectedAt) - Date.parse(compass.collectedAt),
    ) <= maxPairAgeMs;
    const ready = complete && samePeriod && closeInTime;
    let message = '两个页面采集完成后可下载总 Excel。';
    if (complete && !samePeriod) message = '两个页面的采集日期不一致，请用相同日期重新采集。';
    else if (complete && !closeInTime) message = '两个页面的采集时间相隔超过 24 小时，请重新采集。';
    else if (ready) message = '两个页面已采集，请确认属于同一店铺后下载总 Excel。';
    return {
      qianchuan,
      compass,
      ready,
      samePeriod,
      message,
      period: ready ? { startDate: qianchuan.startDate, endDate: qianchuan.endDate } : null,
      items: [statusItem('qianchuan', qianchuan), statusItem('compass', compass)],
    };
  }

  function setNumberFormat(sheet, column, rowCount, format) {
    for (let row = 1; row <= rowCount; row++) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      if (sheet[address]) sheet[address].z = format;
    }
  }

  function buildWorkbook(state = summary()) {
    if (!state.ready) throw new Error(state.message);
    if (typeof XLSX === 'undefined' || !XLSX?.utils) {
      throw new Error('Excel 组件未加载，请刷新页面后重试。');
    }

    const qianchuanRows = [
      ['推广类型', '素材 ID', '素材名称', '消耗（元）', '展现（次）'],
      ...state.qianchuan.rows.map((row) => [
        row.promotionType,
        row.materialId,
        row.title,
        row.spendYuan,
        row.impressions,
      ]),
    ];
    const compassRows = [
      ['标题', '观看次数'],
      ...state.compass.rows.map((row) => [row.title, row.watchCount]),
    ];
    const workbook = XLSX.utils.book_new();
    const qianchuanSheet = XLSX.utils.aoa_to_sheet(qianchuanRows);
    qianchuanSheet['!cols'] = [
      { wch: 12 }, { wch: 22 }, { wch: 48 }, { wch: 16 }, { wch: 16 },
    ];
    qianchuanSheet['!autofilter'] = { ref: qianchuanSheet['!ref'] || 'A1:E1' };
    setNumberFormat(qianchuanSheet, 3, state.qianchuan.rows.length, '#,##0.00');
    setNumberFormat(qianchuanSheet, 4, state.qianchuan.rows.length, '#,##0');
    XLSX.utils.book_append_sheet(workbook, qianchuanSheet, '千川素材');

    const compassSheet = XLSX.utils.aoa_to_sheet(compassRows);
    compassSheet['!cols'] = [{ wch: 58 }, { wch: 18 }];
    compassSheet['!autofilter'] = { ref: compassSheet['!ref'] || 'A1:B1' };
    setNumberFormat(compassSheet, 1, state.compass.rows.length, '#,##0');
    XLSX.utils.book_append_sheet(workbook, compassSheet, '自营视频');
    return workbook;
  }

  function downloadCombined() {
    const state = summary();
    const workbook = buildWorkbook(state);
    const filename = `video-collection-summary-${state.period.startDate}_${state.period.endDate}.xlsx`;
    XLSX.writeFileXLSX(workbook, filename, { compression: true });
    return filename;
  }

  function subscribe(callback) {
    if (typeof GM_addValueChangeListener !== 'function') return [];
    const listeners = [];
    for (const key of Object.values(storageKeys)) {
      try {
        listeners.push(GM_addValueChangeListener(key, callback));
      } catch {
        // Cross-tab status updates are optional when the manager does not support listeners.
      }
    }
    return listeners;
  }

  function unsubscribe(listeners) {
    if (typeof GM_removeValueChangeListener !== 'function') return;
    for (const listener of listeners ?? []) {
      try {
        GM_removeValueChangeListener(listener);
      } catch {
        // The listener may already have been removed by the userscript manager.
      }
    }
  }

  return {
    begin, buildWorkbook, downloadCombined, invalidate, read, save, storageKeys,
    subscribe, summary, unsubscribe,
  };
})();

const videoCollectionLiquidUi = (() => {
  'use strict';

  const themeKey = 'video-collection-liquid-theme';
  // Older installs also saved the automatic bottom-right coordinates.
  const positionKeyPrefix = 'video-collection-liquid-position:v2';
  const edgeGap = 12;
  const snapDistance = 36;
  const liquidCss = `
    :host {
      font-family: "SF Pro Text", "SF Pro Display", "Segoe UI Variable", "PingFang SC",
        "Microsoft YaHei", sans-serif;
      font-size: 13px;
      line-height: 1.45;
    }

    *, *::before, *::after { box-sizing: border-box; }
    [hidden] { display: none !important; }

    .vc-shell {
      --vc-accent: #2563eb;
      --vc-accent-strong: #1d4ed8;
      --vc-accent-soft: rgb(37 99 235 / 12%);
      --vc-glass: rgb(255 255 255 / 66%);
      --vc-glass-strong: rgb(255 255 255 / 92%);
      --vc-standard: rgb(255 255 255 / 84%);
      --vc-control: rgb(241 245 249 / 82%);
      --vc-control-hover: rgb(255 255 255 / 96%);
      --vc-border: rgb(15 23 42 / 7%);
      --vc-border-soft: rgb(15 23 42 / 7%);
      --vc-text: #0f172a;
      --vc-muted: #64748b;
      --vc-warning: #92400e;
      --vc-warning-bg: rgb(254 243 199 / 78%);
      --vc-shadow: 0 24px 60px -12px rgb(15 23 42 / 18%), 0 4px 16px rgb(15 23 42 / 6%);
      --vc-card-shadow: 0 2px 10px rgb(15 23 42 / 3.5%);
      --el-color-primary: var(--vc-accent);
      --el-color-primary-dark-2: var(--vc-accent-strong);
      --el-text-color-primary: var(--vc-text);
      --el-text-color-regular: var(--vc-text);
      --el-text-color-secondary: var(--vc-muted);
      --el-border-color: var(--vc-border-soft);
      --el-border-color-light: var(--vc-border-soft);
      --el-border-color-lighter: rgb(15 23 42 / 5%);
      --el-fill-color-blank: transparent;
      --el-fill-color-light: rgb(241 245 249 / 70%);
      --el-bg-color: var(--vc-standard);
      --el-bg-color-overlay: var(--vc-glass-strong);
      color: var(--vc-text);
      color-scheme: light;
      font-family: inherit;
    }

    .vc-shell.dark {
      --vc-accent: #3b82f6;
      --vc-accent-strong: #60a5fa;
      --vc-accent-soft: rgb(59 130 246 / 18%);
      --vc-glass: rgb(17 24 39 / 72%);
      --vc-glass-strong: rgb(30 41 59 / 94%);
      --vc-standard: rgb(30 41 59 / 82%);
      --vc-control: rgb(15 23 42 / 68%);
      --vc-control-hover: rgb(255 255 255 / 14%);
      --vc-border: rgb(255 255 255 / 9%);
      --vc-border-soft: rgb(255 255 255 / 8%);
      --vc-text: #f8fafc;
      --vc-muted: #94a3b8;
      --vc-warning: #fcd34d;
      --vc-warning-bg: rgb(120 53 15 / 55%);
      --vc-shadow: 0 28px 70px -10px rgb(0 0 0 / 55%), 0 2px 14px rgb(0 0 0 / 32%);
      --vc-card-shadow: 0 4px 14px rgb(0 0 0 / 20%);
      color-scheme: dark;
    }

    .vc-launcher,
    .vc-panel {
      --vc-floating-gap: 20px;
      --vc-default-top: clamp(
        var(--vc-floating-gap),
        var(--vc-preferred-top, 200px),
        calc(100dvh - var(--vc-floating-gap) - 220px)
      );
      position: fixed;
      z-index: 2147483600;
      top: var(--vc-default-top);
      right: var(--vc-floating-gap);
    }

    .vc-shell .vc-launcher.el-button {
      min-height: 46px;
      padding: 0 18px;
      color: var(--vc-text);
      background: var(--vc-glass);
      border: 1px solid var(--vc-border);
      border-radius: 999px;
      box-shadow: var(--vc-shadow);
      backdrop-filter: blur(28px) saturate(180%);
      -webkit-backdrop-filter: blur(28px) saturate(180%);
      font-weight: 700;
      letter-spacing: .01em;
      overflow: hidden;
      cursor: grab;
      touch-action: none;
      transition: transform 120ms ease-out, box-shadow 220ms ease, background 220ms ease;
    }

    .vc-shell .vc-launcher.el-button:hover {
      color: var(--vc-text);
      background: var(--vc-glass-strong);
      transform: translateY(-2px);
    }

    .vc-shell .vc-launcher.el-button:active {
      transform: scale(.97);
    }

    .vc-panel {
      display: flex;
      flex-direction: column;
      width: min(460px, calc(100vw - 40px));
      max-height: min(720px, calc(100dvh - var(--vc-default-top) - var(--vc-floating-gap)));
      overflow: hidden;
      isolation: isolate;
      background: var(--vc-glass);
      border: 1px solid var(--vc-border);
      border-radius: 26px;
      box-shadow: var(--vc-shadow);
      backdrop-filter: blur(28px) saturate(180%);
      -webkit-backdrop-filter: blur(28px) saturate(180%);
    }

    .vc-refraction {
      position: absolute;
      z-index: -2;
      inset: -30% -20% auto auto;
      width: 330px;
      aspect-ratio: 1;
      background:
        radial-gradient(circle at 32% 36%, rgb(255 212 104 / 26%), transparent 38%),
        radial-gradient(circle at 66% 58%, rgb(58 166 255 / 26%), transparent 46%);
      filter: blur(18px);
      opacity: .72;
      pointer-events: none;
      transform: rotate(-14deg);
    }

    .dark .vc-refraction {
      background:
        radial-gradient(circle at 32% 36%, rgb(70 157 255 / 22%), transparent 38%),
        radial-gradient(circle at 65% 58%, rgb(52 229 201 / 12%), transparent 46%);
      opacity: .52;
    }

    .vc-header {
      display: flex;
      flex: 0 0 auto;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      padding: 15px 16px 10px 18px;
      cursor: grab;
      touch-action: none;
      user-select: none;
    }

    .vc-launcher.is-dragging,
    .vc-launcher.is-dragging:hover,
    .vc-panel.is-dragging {
      cursor: grabbing;
      transform: none;
      transition: none;
    }

    .vc-panel.is-dragging .vc-header { cursor: grabbing; }

    .vc-title-wrap {
      display: grid;
      grid-template-columns: auto 1fr;
      align-items: center;
      gap: 10px;
      min-width: 0;
    }

    .vc-title-mark {
      width: 6px;
      height: 20px;
      background: linear-gradient(180deg, #60a5fa, var(--vc-accent));
      border-radius: 999px;
      box-shadow: 0 0 12px rgb(37 99 235 / 35%);
    }

    .vc-header h2 {
      margin: 0;
      overflow: hidden;
      font-size: 16px;
      font-weight: 700;
      letter-spacing: -.01em;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .vc-header-actions {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .vc-shell .el-segmented {
      --el-segmented-bg-color: var(--vc-control);
      --el-segmented-item-selected-bg-color: var(--vc-glass-strong);
      --el-segmented-item-selected-color: var(--vc-text);
      --el-border-radius-base: 999px;
      padding: 2.5px;
      border: 1px solid var(--vc-border-soft);
      border-radius: 999px;
    }

    .vc-shell .el-segmented__item {
      min-width: 46px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      transition: transform 100ms ease-out;
    }

    .vc-shell .el-segmented__item:active {
      transform: scale(.96);
    }

    .vc-shell .vc-close.el-button {
      width: 30px;
      height: 30px;
      min-height: 30px;
      padding: 0;
      color: var(--vc-muted);
      background: var(--vc-glass-strong);
      border: 1px solid var(--vc-border-soft);
      font-size: 18px;
      font-weight: 400;
    }

    .vc-panel-body {
      display: flex;
      flex-direction: column;
      gap: 11px;
      min-height: 0;
      padding: 6px 16px 12px;
      overflow: auto;
      overscroll-behavior: contain;
      scrollbar-width: thin;
      scrollbar-color: var(--vc-border-soft) transparent;
    }

    .vc-panel-body > * { flex-shrink: 0; }

    /* Merged Status Banner */
    .vc-banner {
      display: flex;
      align-items: flex-start;
      gap: 10px;
      padding: 11px 14px;
      border: 1px solid var(--vc-border-soft);
      border-radius: 16px;
      box-shadow: var(--vc-card-shadow);
      transition: background 0.2s ease, color 0.2s ease;
    }

    .vc-banner.waiting {
      background: rgb(254 243 199 / 68%);
      color: #92400e;
    }
    .dark .vc-banner.waiting {
      background: rgb(245 158 11 / 16%);
      color: #fcd34d;
    }

    .vc-banner.ready,
    .vc-banner.done {
      background: rgb(209 250 229 / 68%);
      color: #065f46;
    }
    .dark .vc-banner.ready,
    .dark .vc-banner.done {
      background: rgb(16 185 129 / 18%);
      color: #6ee7b7;
    }

    .vc-banner.collecting {
      background: rgb(219 234 254 / 75%);
      color: #1e40af;
    }
    .dark .vc-banner.collecting {
      background: rgb(59 130 246 / 22%);
      color: #93c5fd;
    }

    .vc-banner-dot {
      position: relative;
      flex-shrink: 0;
      width: 8px;
      height: 8px;
      margin-top: 5px;
      background: currentColor;
      border-radius: 50%;
    }

    .vc-banner.waiting .vc-banner-dot::after,
    .vc-banner.collecting .vc-banner-dot::after {
      content: "";
      position: absolute;
      inset: -3px;
      border: 1.5px solid currentColor;
      border-radius: 50%;
      opacity: .55;
      animation: vc-ping 1.6s cubic-bezier(0, 0, .2, 1) infinite;
    }

    @keyframes vc-ping {
      75%, 100% { transform: scale(1.85); opacity: 0; }
    }

    .vc-banner-copy { min-width: 0; flex: 1; }
    .vc-banner-title { font-size: 13px; font-weight: 700; line-height: 1.35; }
    .vc-banner-pills {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-bottom: 4px;
    }
    .vc-banner-pill {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 2px 8px;
      background: rgb(255 255 255 / 55%);
      border-radius: 999px;
      font-size: 11.5px;
      font-weight: 600;
      color: var(--vc-text);
    }
    .dark .vc-banner-pill { background: rgb(255 255 255 / 10%); }
    .vc-banner-pill[data-state="ready"] .vc-state-dot,
    .vc-banner-pill[data-state="done"] .vc-state-dot { color: #10b981; }
    .vc-banner-pill[data-state="waiting"] .vc-state-dot { color: #94a3b8; }
    .vc-banner-pill[data-state="active"] .vc-state-dot { color: var(--vc-accent); }
    .vc-banner-desc {
      margin-top: 2px;
      font-size: 11.5px;
      letter-spacing: .01em;
      line-height: 1.4;
      opacity: .9;
      overflow-wrap: anywhere;
    }
    .vc-banner-note {
      margin-top: 5px;
      padding-top: 5px;
      border-top: 1px solid rgb(15 23 42 / 8%);
      font-size: 11px;
      line-height: 1.4;
      opacity: .85;
      overflow-wrap: anywhere;
    }
    .dark .vc-banner-note { border-top-color: rgb(255 255 255 / 10%); }

    /* Inner Card */
    .vc-card {
      background: var(--vc-standard);
      border: 1px solid var(--vc-border-soft);
      border-radius: 18px;
      padding: 12px 14px;
      box-shadow: var(--vc-card-shadow);
    }

    .vc-card-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      margin-bottom: 10px;
    }

    .vc-card-title {
      color: var(--vc-muted);
      font-size: 12.5px;
      font-weight: 600;
    }

    .vc-card-hint {
      color: var(--vc-muted);
      font-size: 11.5px;
      letter-spacing: .01em;
      text-align: right;
    }

    .vc-preset-pills {
      display: flex;
      gap: 5px;
    }

    .vc-preset-pill {
      padding: 4px 11px;
      color: var(--vc-muted);
      background: var(--vc-control);
      border: 1px solid transparent;
      border-radius: 999px;
      font-family: inherit;
      font-size: 11.5px;
      font-weight: 600;
      cursor: pointer;
      transition: transform 100ms ease-out, background 160ms ease, color 160ms ease;
    }

    .vc-preset-pill:hover:not(:disabled) {
      color: var(--vc-text);
      background: var(--vc-glass-strong);
    }

    .vc-preset-pill:active:not(:disabled) { transform: scale(.96); }

    .vc-preset-pill.active {
      color: var(--vc-accent);
      background: var(--vc-glass-strong);
      border-color: var(--vc-border-soft);
      box-shadow: 0 1px 4px rgb(37 99 235 / 12%);
    }

    .vc-preset-pill:disabled { opacity: .55; cursor: not-allowed; }

    .vc-shell .el-button {
      --el-button-text-color: var(--vc-text);
      --el-button-bg-color: var(--vc-glass-strong);
      --el-button-border-color: var(--vc-border-soft);
      --el-button-hover-text-color: var(--vc-text);
      --el-button-hover-bg-color: var(--vc-control-hover);
      --el-button-hover-border-color: var(--vc-border);
      min-height: 38px;
      border-radius: 999px;
      font-weight: 600;
      transition: transform 100ms ease-out, background 160ms ease, opacity 160ms ease;
    }

    .vc-shell .el-button:not(.is-disabled):active { transform: scale(.97); }

    .vc-shell .el-button--primary {
      --el-button-text-color: #fff;
      --el-button-bg-color: var(--vc-accent);
      --el-button-border-color: transparent;
      --el-button-hover-text-color: #fff;
      --el-button-hover-bg-color: var(--vc-accent-strong);
      --el-button-hover-border-color: transparent;
      background: linear-gradient(180deg, #3b82f6 0%, #2563eb 100%);
      box-shadow: 0 4px 12px rgb(37 99 235 / 26%);
    }

    .vc-shell .el-button--primary.is-disabled,
    .vc-shell .el-button--primary.is-disabled:hover {
      color: var(--vc-muted);
      background: var(--vc-control);
      border-color: var(--vc-border-soft);
      box-shadow: none;
      opacity: .85;
    }

    .vc-date-range {
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      gap: 8px;
      padding: 6px 12px;
      background: var(--vc-control);
      border: 1px solid var(--vc-border-soft);
      border-radius: 13px;
    }

    .vc-date-range:focus-within {
      border-color: var(--vc-accent);
      box-shadow: 0 0 0 3px var(--vc-accent-soft);
    }

    .vc-date-input {
      width: 100%;
      min-width: 0;
      height: 26px;
      padding: 0;
      color: var(--vc-text);
      background: transparent;
      border: none;
      font: inherit;
      font-size: 12.5px;
      font-weight: 600;
      font-variant-numeric: tabular-nums;
      outline: none;
    }

    .vc-date-input:disabled { opacity: .56; }
    .dark .vc-date-input::-webkit-calendar-picker-indicator { filter: invert(1); opacity: .75; }
    .vc-date-sep { color: var(--vc-muted); font-size: 12px; }

    .vc-scope-note {
      padding: 10px 12px;
      color: var(--vc-warning);
      background: var(--vc-warning-bg);
      border: 1px solid rgb(222 163 53 / 22%);
      border-radius: 14px;
      font-size: 12px;
      line-height: 1.45;
      overflow-wrap: anywhere;
    }

    .vc-state-dot {
      width: 6px;
      height: 6px;
      background: currentColor;
      border-radius: 50%;
    }

    .vc-workbook-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
    }

    .vc-workbook-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      min-width: 0;
      padding: 10px 12px;
      background: var(--vc-control);
      border: 1px solid var(--vc-border-soft);
      border-radius: 13px;
    }

    .vc-workbook-name {
      font-size: 12.5px;
      font-weight: 600;
      color: var(--vc-text);
    }

    .vc-workbook-badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      color: var(--vc-muted);
      font-size: 11.5px;
      font-weight: 600;
      letter-spacing: .01em;
      white-space: nowrap;
    }

    .vc-workbook-item[data-state="done"] .vc-workbook-badge { color: #059669; }
    .dark .vc-workbook-item[data-state="done"] .vc-workbook-badge { color: #34d399; }
    .vc-workbook-item[data-state="active"] .vc-workbook-badge { color: var(--vc-accent); }

    .vc-filter-list {
      display: flex;
      flex-wrap: wrap;
      gap: 5px;
      justify-content: flex-end;
    }

    .vc-shell .el-tag {
      --el-tag-bg-color: var(--vc-control);
      --el-tag-border-color: var(--vc-border-soft);
      --el-tag-text-color: var(--vc-muted);
      border-radius: 999px;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: .01em;
    }

    /* KPI Metrics */
    .vc-kpi-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }

    .vc-kpi-grid.is-quad {
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 8px;
    }

    .vc-kpi-box {
      padding: 10px 12px;
      background: var(--vc-control);
      border: 1px solid var(--vc-border-soft);
      border-radius: 13px;
    }

    .vc-kpi-grid.is-quad .vc-kpi-box {
      padding: 9px 10px;
    }

    .vc-kpi-label {
      margin-bottom: 2px;
      color: var(--vc-muted);
      font-size: 11.5px;
      letter-spacing: .01em;
    }

    .vc-kpi-value {
      color: var(--vc-text);
      font-size: 21px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
      letter-spacing: -.02em;
      line-height: 1.2;
    }

    .vc-kpi-grid.is-quad .vc-kpi-value {
      font-size: 17px;
    }

    .vc-kpi-box.highlight .vc-kpi-value { color: var(--vc-accent); }

    .vc-kpi-unit {
      margin-left: 2px;
      color: var(--vc-muted);
      font-size: 12px;
      font-weight: 500;
    }

    .vc-progress-track {
      height: 5px;
      margin-top: 10px;
      overflow: hidden;
      background: var(--vc-control);
      border-radius: 999px;
    }

    .vc-progress-fill {
      height: 100%;
      background: linear-gradient(90deg, #60a5fa, #2563eb);
      border-radius: 999px;
      transition: width 250ms ease;
    }

    .vc-table-scroll {
      max-height: 210px;
      margin-top: 10px;
      overflow: auto;
      background: transparent;
      border: 1px solid var(--vc-border-soft);
      border-radius: 13px;
    }

    /* Sticky Action Footer */
    .vc-footer {
      display: flex;
      flex: 0 0 auto;
      align-items: center;
      gap: 8px;
      padding: 8px 16px 15px;
    }

    .vc-footer .el-button {
      flex: 1;
      margin-left: 0 !important;
      font-size: 12.5px;
    }

    .vc-footer .el-button--primary {
      flex: 1.25;
      font-size: 13px;
    }

    .vc-result-table {
      width: 100%;
      color: var(--vc-text);
      background: transparent;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 13px;
    }

    .vc-result-table th,
    .vc-result-table td {
      padding: 9px 11px;
      border-bottom: 1px solid var(--vc-border-soft);
      line-height: 1.45;
      text-align: left;
      vertical-align: top;
      overflow-wrap: anywhere;
    }

    .vc-result-table th {
      position: sticky;
      z-index: 1;
      top: 0;
      color: var(--vc-muted);
      background: var(--vc-control);
      font-size: 12px;
      font-weight: 720;
    }

    .vc-result-table tbody tr:last-child td { border-bottom: 0; }
    .vc-result-table tbody tr:hover td { background: var(--vc-accent-soft); }
    .vc-result-table .vc-number { text-align: right; }
    .vc-number { font-variant-numeric: tabular-nums; white-space: nowrap; }

    .vc-panel-enter-active,
    .vc-panel-leave-active { transition: opacity 200ms ease, transform 240ms cubic-bezier(.2, .8, .2, 1); }
    .vc-panel-enter-from,
    .vc-panel-leave-to { opacity: 0; transform: translateY(12px) scale(.975); }
    .vc-launcher-enter-active,
    .vc-launcher-leave-active { transition: opacity 160ms ease, transform 180ms ease; }
    .vc-launcher-enter-from,
    .vc-launcher-leave-to { opacity: 0; transform: translateY(6px) scale(.94); }

    @keyframes vc-pulse {
      0%, 100% { transform: scale(.9); opacity: .72; }
      50% { transform: scale(1.08); opacity: 1; }
    }

    @supports not ((backdrop-filter: blur(1px)) or (-webkit-backdrop-filter: blur(1px))) {
      .vc-panel, .vc-shell .vc-launcher.el-button { background: var(--vc-glass-strong); }
    }

    @media (max-width: 560px) {
      .vc-launcher, .vc-panel { --vc-floating-gap: 12px; }
      .vc-panel { width: calc(100vw - 24px); border-radius: 24px; }
      .vc-panel-body { padding-inline: 13px; }
      .vc-header { padding-inline: 15px 11px; }
      .vc-workbook-grid { grid-template-columns: 1fr; }
      .vc-kpi-grid.is-quad { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .vc-shell .el-segmented__item { min-width: 42px; }
      .vc-result-table { table-layout: auto; font-size: 12px; }
      .vc-result-table col { width: auto !important; min-width: 0 !important; }
      .vc-result-table th, .vc-result-table td { padding-inline: 8px; }
    }

    @media (prefers-reduced-motion: reduce) {
      .vc-shell *, .vc-shell *::before, .vc-shell *::after {
        scroll-behavior: auto !important;
        animation-duration: .01ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: .01ms !important;
      }
    }

    @media (prefers-reduced-transparency: reduce) {
      .vc-panel, .vc-shell .vc-launcher.el-button {
        backdrop-filter: none !important;
        -webkit-backdrop-filter: none !important;
        background: var(--vc-glass-strong);
      }
    }

    @media (prefers-contrast: more) {
      .vc-shell {
        --vc-glass: #f8fafc;
        --vc-standard: #ffffff;
        --vc-border: rgb(15 23 42 / 28%);
        --vc-border-soft: rgb(15 23 42 / 24%);
      }
      .vc-shell.dark {
        --vc-glass: #0b1120;
        --vc-standard: #1e293b;
        --vc-border: rgb(255 255 255 / 35%);
        --vc-border-soft: rgb(255 255 255 / 28%);
      }
    }
  `;

  const template = `
    <div class="vc-shell" :class="{ dark: theme === 'dark' }" :data-theme="theme" @keydown.esc="closePanel">
      <transition name="vc-launcher">
        <el-button
          v-show="!panelOpen"
          ref="launcher"
          class="vc-launcher"
          :class="{ 'is-dragging': dragKind === 'launcher' }"
          :style="floatingStyle('launcher')"
          round
          :aria-expanded="String(panelOpen)"
          @pointerdown="startDrag('launcher', $event)"
          @click="openPanel"
        >
          {{ config.launcherLabel }}
        </el-button>
      </transition>

      <transition name="vc-panel">
        <section
          v-show="panelOpen"
          ref="panel"
          class="vc-panel"
          :class="{ 'is-dragging': dragKind === 'panel' }"
          :style="[{ '--vc-preferred-top': config.panelTop }, floatingStyle('panel')]"
          role="dialog"
          aria-modal="false"
          :aria-label="config.title"
        >
          <div class="vc-refraction" aria-hidden="true"></div>
          <header class="vc-header" @pointerdown="startDrag('panel', $event)">
            <div class="vc-title-wrap">
              <span class="vc-title-mark" aria-hidden="true"></span>
              <h2>{{ config.title }}</h2>
            </div>
            <div class="vc-header-actions">
              <el-segmented
                v-model="theme"
                :options="themeOptions"
                size="small"
                aria-label="主题"
                @change="saveTheme"
              ></el-segmented>
              <el-button
                ref="closeButton"
                class="vc-close"
                circle
                text
                aria-label="关闭面板"
                @click="closePanel"
              >×</el-button>
            </div>
          </header>

          <div class="vc-panel-body">
            <div class="vc-banner" :class="banner.state" role="status">
              <span class="vc-banner-dot" aria-hidden="true"></span>
              <div class="vc-banner-copy">
                <div v-if="view.progress && view.progress.length > 1" class="vc-banner-pills">
                  <span
                    v-for="item in view.progress"
                    :key="item.key"
                    class="vc-banner-pill"
                    :data-state="item.state"
                  >
                    <span class="vc-state-dot" aria-hidden="true"></span>
                    <span>{{ item.text }}</span>
                  </span>
                </div>
                <div v-else class="vc-banner-title">{{ banner.title }}</div>
                <div v-if="banner.desc" class="vc-banner-desc">{{ banner.desc }}</div>
                <div v-if="view.scopeNote" class="vc-banner-note">{{ view.scopeNote }}</div>
              </div>
            </div>

            <section class="vc-card">
              <div class="vc-card-head">
                <span class="vc-card-title">{{ view.periodLabel }}范围</span>
                <div class="vc-preset-pills">
                  <button
                    type="button"
                    class="vc-preset-pill"
                    :class="{ active: activePreset === '7d' }"
                    :disabled="view.running"
                    @click="useRecentDays(7)"
                  >近7天</button>
                  <button
                    type="button"
                    class="vc-preset-pill"
                    :class="{ active: activePreset === 'month' }"
                    :disabled="view.running"
                    @click="useMonth"
                  >本月</button>
                </div>
              </div>
              <div class="vc-date-range">
                <input
                  v-model="dateStart"
                  class="vc-date-input"
                  type="date"
                  aria-label="开始日期"
                  required
                  :disabled="view.running"
                  @change="applyDates"
                >
                <span class="vc-date-sep">至</span>
                <input
                  v-model="dateEnd"
                  class="vc-date-input"
                  type="date"
                  aria-label="结束日期"
                  required
                  :disabled="view.running"
                  @change="applyDates"
                >
              </div>
            </section>

            <section class="vc-card" aria-label="汇总进度">
              <div class="vc-card-head">
                <span class="vc-card-title">汇总进度</span>
                <span class="vc-card-hint" role="status">{{ workbookError || view.workbookStatus }}</span>
              </div>
              <div class="vc-workbook-grid">
                <div
                  v-for="item in view.workbookItems"
                  :key="item.key"
                  class="vc-workbook-item"
                  :data-state="item.state"
                >
                  <span class="vc-workbook-name">{{ item.label }}</span>
                  <span class="vc-workbook-badge">
                    <span class="vc-state-dot" aria-hidden="true"></span>
                    <span>{{ item.state === 'done' ? '已完成' : item.state === 'active' ? '采集中' : '未采集' }}</span>
                  </span>
                </div>
              </div>
            </section>

            <section class="vc-card">
              <div class="vc-card-head">
                <span class="vc-card-title">当前页面统计</span>
                <div class="vc-filter-list">
                  <el-tag v-for="filter in view.filters" :key="filter" effect="plain" round>{{ filter }}</el-tag>
                </div>
              </div>
              <div class="vc-kpi-grid" :class="{ 'is-quad': view.kpis && view.kpis.length > 2 }">
                <div
                  v-for="kpi in view.kpis"
                  :key="kpi.label"
                  class="vc-kpi-box"
                  :class="{ highlight: kpi.highlight }"
                >
                  <div class="vc-kpi-label">{{ kpi.label }}</div>
                  <div class="vc-kpi-value">{{ kpi.value }}<span class="vc-kpi-unit">{{ kpi.unit }}</span></div>
                </div>
              </div>
              <div v-if="view.progressPercent > 0" class="vc-progress-track">
                <div class="vc-progress-fill" :style="{ width: view.progressPercent + '%' }"></div>
              </div>
              <div v-if="view.rows.length" class="vc-table-scroll">
                <table class="vc-result-table" :aria-label="config.tableLabel">
                  <colgroup>
                    <col
                      v-for="column in view.columns"
                      :key="column.prop"
                      :style="{
                        width: column.width ? column.width + 'px' : undefined,
                        minWidth: column.minWidth ? column.minWidth + 'px' : undefined,
                      }"
                    >
                  </colgroup>
                  <thead>
                    <tr>
                      <th
                        v-for="column in view.columns"
                        :key="column.prop"
                        scope="col"
                        :class="{ 'vc-number': column.numeric }"
                      >{{ column.label }}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="row in view.rows" :key="row.key">
                      <td
                        v-for="column in view.columns"
                        :key="column.prop"
                        :class="{ 'vc-number': column.numeric }"
                      >{{ row[column.prop] }}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
          </div>

          <footer class="vc-footer">
            <el-button
              v-if="!view.running"
              ref="startButton"
              type="primary"
              :disabled="view.startDisabled"
              @click="startRun"
            >{{ view.startLabel }}</el-button>
            <el-button
              v-else
              type="primary"
              :disabled="view.pauseDisabled"
              @click="pauseRun"
            >{{ view.pauseDisabled ? '正在暂停' : '暂停采集' }}</el-button>
            <el-button :disabled="view.downloadDisabled" @click="download">导出 CSV</el-button>
            <el-button
              :disabled="view.combinedDownloadDisabled"
              @click="downloadCombined"
            >总 Excel</el-button>
          </footer>
        </section>
      </transition>
    </div>
  `;

  function resourceText(name) {
    try {
      return typeof GM_getResourceText === 'function' ? GM_getResourceText(name) : '';
    } catch {
      return '';
    }
  }

  function loadTheme() {
    try {
      if (typeof GM_getValue === 'function') {
        return GM_getValue(themeKey, 'sun') === 'dark' ? 'dark' : 'sun';
      }
    } catch {
      // Fall through to origin-local storage when userscript storage is unavailable.
    }
    try {
      return localStorage.getItem(themeKey) === 'dark' ? 'dark' : 'sun';
    } catch {
      return 'sun';
    }
  }

  function storeTheme(theme) {
    const value = theme === 'dark' ? 'dark' : 'sun';
    try {
      if (typeof GM_setValue === 'function') {
        GM_setValue(themeKey, value);
        return;
      }
    } catch {
      // Fall through to origin-local storage when userscript storage is unavailable.
    }
    try {
      localStorage.setItem(themeKey, value);
    } catch {
      // Theme persistence is optional when storage is unavailable.
    }
  }

  function floatingPositionKey(hostId, kind) {
    return `${positionKeyPrefix}:${hostId}:${kind}`;
  }

  function validPosition(value) {
    if (!value || !Number.isFinite(Number(value.x)) || !Number.isFinite(Number(value.y))) {
      return null;
    }
    return { x: Number(value.x), y: Number(value.y) };
  }

  function loadFloatingPosition(key) {
    try {
      if (typeof GM_getValue === 'function') return validPosition(GM_getValue(key, null));
    } catch {
      // Fall through to origin-local storage when userscript storage is unavailable.
    }
    try {
      return validPosition(JSON.parse(localStorage.getItem(key) ?? 'null'));
    } catch {
      return null;
    }
  }

  function storeFloatingPosition(key, position) {
    const value = validPosition(position);
    if (!value) return;
    try {
      if (typeof GM_setValue === 'function') {
        GM_setValue(key, value);
        return;
      }
    } catch {
      // Fall through to origin-local storage when userscript storage is unavailable.
    }
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // Position persistence is optional when storage is unavailable.
    }
  }

  function fitFloatingPosition(position, size, viewport, gap = edgeGap, threshold = snapDistance) {
    const width = Math.max(0, Number(size?.width) || 0);
    const height = Math.max(0, Number(size?.height) || 0);
    const viewportWidth = Math.max(width + gap * 2, Number(viewport?.width) || 0);
    const viewportHeight = Math.max(height + gap * 2, Number(viewport?.height) || 0);
    const maxX = Math.max(gap, viewportWidth - width - gap);
    const maxY = Math.max(gap, viewportHeight - height - gap);
    let x = Math.min(maxX, Math.max(gap, Number(position?.x) || gap));
    let y = Math.min(maxY, Math.max(gap, Number(position?.y) || gap));

    if (threshold >= 0) {
      const leftDistance = x - gap;
      const rightDistance = maxX - x;
      const topDistance = y - gap;
      const bottomDistance = maxY - y;
      if (Math.min(leftDistance, rightDistance) <= threshold) {
        x = leftDistance <= rightDistance ? gap : maxX;
      }
      if (Math.min(topDistance, bottomDistance) <= threshold) {
        y = topDistance <= bottomDistance ? gap : maxY;
      }
    }

    return { x: Math.round(x), y: Math.round(y) };
  }

  function mount(options) {
    const host = document.createElement('div');
    host.id = options.hostId;
    const shadow = host.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    const elementCss = resourceText('ELEMENT_PLUS_CSS').replaceAll(':root', ':host');
    const darkCss = resourceText('ELEMENT_PLUS_DARK_CSS').replaceAll('html.dark', '.vc-shell.dark');
    style.textContent = elementCss + '\n' + darkCss + '\n' + liquidCss;
    const root = document.createElement('div');
    shadow.append(style, root);
    document.body.append(host);

    if (typeof Vue === 'undefined' || typeof ElementPlus === 'undefined') {
      const fallback = document.createElement('button');
      fallback.type = 'button';
      fallback.className = 'vc-launcher';
      fallback.textContent = '采集面板资源加载失败，请刷新页面';
      fallback.style.cssText = 'padding:12px 16px;border:1px solid #ccd5e0;border-radius:16px;background:#fff;color:#1d293a;';
      root.append(fallback);
      return {
        open: () => {},
        refresh: () => { host.hidden = !options.isVisible(); },
      };
    }

    const positionKeys = {
      launcher: floatingPositionKey(options.hostId, 'launcher'),
      panel: floatingPositionKey(options.hostId, 'panel'),
    };

    const app = Vue.createApp({
      data() {
        const initial = options.initialPeriod();
        return {
          config: options.config,
          revision: 0,
          panelOpen: false,
          dragKind: '',
          dragState: null,
          ignoreLauncherClick: false,
          storageListeners: [],
          workbookError: '',
          positions: {
            launcher: loadFloatingPosition(positionKeys.launcher),
            panel: loadFloatingPosition(positionKeys.panel),
          },
          theme: loadTheme(),
          themeOptions: [
            { label: '☀️ 浅色', value: 'sun' },
            { label: '🌙 深色', value: 'dark' },
          ],
          activePreset: 'month',
          dateStart: initial.startDate,
          dateEnd: initial.endDate,
        };
      },
      computed: {
        view() {
          this.revision;
          return options.getView();
        },
        banner() {
          const v = this.view;
          const progressText = (v.progress || []).map((p) => p.text).join(' ');
          if (v.running) {
            return { state: 'collecting', title: progressText || '正在采集中...', desc: v.status };
          }
          if (!v.combinedDownloadDisabled) {
            return { state: 'done', title: '全部采集完成', desc: v.workbookStatus || v.status };
          }
          if (!v.downloadDisabled) {
            return { state: 'done', title: progressText || '当前页面采集完成', desc: v.status };
          }
          if (!v.startDisabled) {
            return { state: 'ready', title: progressText || '接口已捕获 随时可以采集', desc: v.status || '点击下方「开始采集」自动拉取当前日期范围数据' };
          }
          return { state: 'waiting', title: progressText || '等待捕获列表请求', desc: v.status };
        },
      },
      methods: {
        floatingElement(kind) {
          const reference = this.$refs[kind];
          return reference?.$el ?? reference ?? null;
        },
        floatingStyle(kind) {
          const position = this.positions[kind];
          if (!position) return undefined;
          return {
            right: 'auto',
            left: `${position.x}px`,
            top: `${position.y}px`,
          };
        },
        normalizePosition(kind) {
          const position = this.positions[kind];
          const element = this.floatingElement(kind);
          if (!position || !element?.offsetWidth || !element.offsetHeight) return;
          this.positions[kind] = fitFloatingPosition(
            position,
            { width: element.offsetWidth, height: element.offsetHeight },
            { width: window.innerWidth, height: window.innerHeight },
            edgeGap,
            -1,
          );
        },
        startDrag(kind, event) {
          if (event.button !== 0 || event.isPrimary === false) return;
          if (kind === 'panel' && event.target?.closest?.('.vc-header-actions')) return;
          const element = this.floatingElement(kind);
          if (!element) return;
          const rect = element.getBoundingClientRect();
          const previousPosition = this.positions[kind]
            ? { ...this.positions[kind] } : null;
          this.positions[kind] = { x: rect.left, y: rect.top };
          this.dragKind = kind;
          this.dragState = {
            kind,
            element,
            pointerId: event.pointerId,
            pointerX: event.clientX,
            pointerY: event.clientY,
            previousPosition,
            originX: rect.left,
            originY: rect.top,
            width: rect.width,
            height: rect.height,
            moved: false,
          };
          element.setPointerCapture?.(event.pointerId);
          element.addEventListener('lostpointercapture', this.endDrag);
          window.addEventListener('pointermove', this.moveDrag, { passive: false });
          window.addEventListener('pointerup', this.endDrag);
          window.addEventListener('pointercancel', this.endDrag);
        },
        moveDrag(event) {
          const drag = this.dragState;
          if (!drag || event.pointerId !== drag.pointerId) return;
          const deltaX = event.clientX - drag.pointerX;
          const deltaY = event.clientY - drag.pointerY;
          if (!drag.moved && Math.hypot(deltaX, deltaY) <= 4) return;
          drag.moved = true;
          event.preventDefault();
          this.positions[drag.kind] = fitFloatingPosition(
            { x: drag.originX + deltaX, y: drag.originY + deltaY },
            { width: drag.width, height: drag.height },
            { width: window.innerWidth, height: window.innerHeight },
            edgeGap,
            -1,
          );
        },
        endDrag(event) {
          const drag = this.dragState;
          if (!drag || event.pointerId !== drag.pointerId) return;
          drag.element?.removeEventListener('lostpointercapture', this.endDrag);
          if (drag.element?.hasPointerCapture?.(drag.pointerId)) {
            drag.element.releasePointerCapture(drag.pointerId);
          }
          window.removeEventListener('pointermove', this.moveDrag);
          window.removeEventListener('pointerup', this.endDrag);
          window.removeEventListener('pointercancel', this.endDrag);
          if (!drag.moved) {
            this.positions[drag.kind] = drag.previousPosition;
            this.dragKind = '';
            this.dragState = null;
            return;
          }
          const position = fitFloatingPosition(
            this.positions[drag.kind],
            { width: drag.width, height: drag.height },
            { width: window.innerWidth, height: window.innerHeight },
          );
          this.positions[drag.kind] = position;
          storeFloatingPosition(positionKeys[drag.kind], position);
          if (drag.kind === 'launcher' && drag.moved) {
            this.ignoreLauncherClick = true;
            setTimeout(() => { this.ignoreLauncherClick = false; }, 0);
          }
          this.dragKind = '';
          this.dragState = null;
        },
        handleViewportResize() {
          this.normalizePosition('launcher');
          if (this.panelOpen) this.normalizePosition('panel');
        },
        sync() {
          this.revision++;
          const next = options.getView();
          if (next.periodAutomatic && !next.running) {
            this.dateStart = next.startDate;
            this.dateEnd = next.endDate;
          }
          host.hidden = !options.isVisible();
        },
        openPanel() {
          if (this.ignoreLauncherClick) return;
          if (!options.isVisible()) return;
          this.panelOpen = true;
          this.sync();
          Vue.nextTick(() => {
            this.normalizePosition('panel');
            const startButton = this.$refs.startButton?.$el;
            const target = startButton && !startButton.disabled
              ? startButton
              : this.$refs.closeButton?.$el;
            target?.focus({ preventScroll: true });
          });
        },
        closePanel() {
          this.panelOpen = false;
          Vue.nextTick(() => {
            this.normalizePosition('launcher');
            this.$refs.launcher?.$el?.focus();
          });
        },
        saveTheme(value) {
          storeTheme(value);
        },
        applyDates() {
          this.activePreset = '';
          options.applyDates(this.dateStart, this.dateEnd, false);
          this.sync();
        },
        useRecentDays(days) {
          const base = this.dateEnd ? new Date(`${this.dateEnd}T00:00:00`) : new Date();
          const start = new Date(base.getTime() - ((days - 1) * 24 * 60 * 60 * 1000));
          const pad = (n) => String(n).padStart(2, '0');
          this.dateStart = `${start.getFullYear()}-${pad(start.getMonth() + 1)}-${pad(start.getDate())}`;
          this.dateEnd = `${base.getFullYear()}-${pad(base.getMonth() + 1)}-${pad(base.getDate())}`;
          this.activePreset = '7d';
          options.applyDates(this.dateStart, this.dateEnd, false);
          this.sync();
        },
        useMonth() {
          const period = options.monthPeriod();
          this.dateStart = period.startDate;
          this.dateEnd = period.endDate;
          this.activePreset = 'month';
          options.applyDates(this.dateStart, this.dateEnd, true);
          this.sync();
        },
        startRun() {
          options.start();
        },
        pauseRun() {
          options.pause();
          this.sync();
        },
        download() {
          options.download();
        },
        downloadCombined() {
          try {
            options.downloadCombined();
            this.workbookError = '';
          } catch (error) {
            this.workbookError = error?.message || '总 Excel 下载失败，请刷新页面后重试。';
          }
        },
      },
      mounted() {
        window.addEventListener('resize', this.handleViewportResize);
        this.storageListeners = options.subscribe?.(() => this.sync()) ?? [];
        Vue.nextTick(() => this.normalizePosition('launcher'));
      },
      beforeUnmount() {
        options.unsubscribe?.(this.storageListeners);
        window.removeEventListener('resize', this.handleViewportResize);
        window.removeEventListener('pointermove', this.moveDrag);
        window.removeEventListener('pointerup', this.endDrag);
        window.removeEventListener('pointercancel', this.endDrag);
      },
      template,
    });
    let vm;
    try {
      app.use(ElementPlus);
      vm = app.mount(root);
    } finally {
      if (typeof __vcRestoreFrameworkGlobals === 'function') __vcRestoreFrameworkGlobals();
    }

    const refresh = () => vm.sync();
    window.addEventListener('popstate', refresh);
    window.navigation?.addEventListener('currententrychange', refresh);
    refresh();
    return { open: () => vm.openPanel(), refresh };
  }

  return { fitFloatingPosition, mount };
})();

(() => {
  'use strict';

  const origin = 'https://qianchuan.jinritemai.com';
  const pagePath = '/dataV3/roi2-material-analysis';
  const statPath = '/ad/api/data/v1/common/statQuery';
  const modeCodes = ['2', '1'];
  const modes = {
    '2': { label: '推直播间', impressionMetric: 'live_show_count_for_roi2_v2' },
    '1': { label: '推商品', impressionMetric: 'product_show_count_for_roi2' },
  };
  const minSpend = 1000;
  const minImpressions = 20000;
  const spendMetric = 'stat_cost_for_roi2';
  const batchLimit = 100;
  const pageDelayMs = 300;
  const requestTimeoutMs = 15000;
  const previewLimit = 50;
  const numberFormat = new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 });

  function currentMonthPeriod(now = new Date()) {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(now);
    const date = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    const yesterday = new Date(Date.UTC(Number(date.year), Number(date.month) - 1, Number(date.day) - 1));
    const endDate = yesterday.toISOString().slice(0, 10);
    return {
      startTime: `${endDate.slice(0, 7)}-01 00:00:00`,
      endTime: `${endDate} 23:59:59`,
    };
  }

  function recentPageMonthPeriod(list) {
    const currentMonth = currentMonthPeriod();
    const defaultEnd = currentMonth.endTime.slice(0, 10);
    const defaultEndTime = Date.parse(`${defaultEnd}T00:00:00Z`);
    const recentEnds = list.filter(Boolean).map((template) => template.endTime.slice(0, 10))
      .filter((endDate) => endDate.slice(0, 7) === defaultEnd.slice(0, 7))
      .filter((endDate) => {
        const lag = (defaultEndTime - Date.parse(`${endDate}T00:00:00Z`)) / 86400000;
        return Number.isInteger(lag) && lag >= 0 && lag <= 7;
      });
    if (!recentEnds.length) return currentMonth;
    return {
      startTime: currentMonth.startTime,
      endTime: `${recentEnds.sort().at(-1)} 23:59:59`,
    };
  }

  function samePeriod(left, right) {
    return left?.startTime === right?.startTime && left?.endTime === right?.endTime;
  }

  function periodText(period) {
    const fullDays = period.startTime.endsWith(' 00:00:00') && period.endTime.endsWith(' 23:59:59');
    return fullDays
      ? `${period.startTime.slice(0, 10)} 至 ${period.endTime.slice(0, 10)}`
      : `${period.startTime} 至 ${period.endTime}`;
  }

  function isDate(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(value) && !value.startsWith('0000')
      && new Date(`${value}T00:00:00Z`).toJSON()?.slice(0, 10) === value;
  }

  function currentPage() {
    if (location.origin !== origin || location.pathname !== pagePath) return null;
    return new URLSearchParams(location.search).get('aavid') ?? '';
  }

  if (currentPage() === null) return;

  let captureAccount = currentPage();
  const templates = new Map();
  let nextSequence = 1;
  let batch;
  let selectedPeriod = currentMonthPeriod();
  let periodAutomatic = true;
  let dateError = '';
  let updatePanel = () => {};
  const requests = new WeakMap();
  const replayRequests = new WeakSet();
  const pageFetch = window.fetch;
  const Xhr = window.XMLHttpRequest;
  const xhrPrototype = Xhr.prototype;
  const xhrOpen = xhrPrototype.open;
  const xhrSetHeader = xhrPrototype.setRequestHeader;
  const xhrSend = xhrPrototype.send;

  function matchEndpoint(method, input) {
    if (String(method).toUpperCase() !== 'POST') return null;
    const url = new URL(input, location.href);
    const source = url.searchParams.get('reqFrom');
    return url.origin === origin && url.pathname === statPath
      && (source === null || source === 'roi2_material_list') ? url : null;
  }

  function singleFilter(conditions, field) {
    const values = conditions.find((item) => item?.Field === field)?.Values;
    return Array.isArray(values) && values.length === 1 ? String(values[0]) : '';
  }

  function readTemplate(request, body, send) {
    let data;
    try {
      data = JSON.parse(body);
    } catch {
      return null;
    }
    if (data?.DataSetKey !== 'roi2_video_material_analysis'
      || (data.reqFrom ?? request.url.searchParams.get('reqFrom')) !== 'roi2_material_list') return null;

    const conditions = data.Filters?.Conditions;
    if (!Array.isArray(conditions) || singleFilter(conditions, 'material_type') !== '3') return null;
    const mode = singleFilter(conditions, 'marketing_goal');
    if (!modeCodes.includes(mode)) return null;

    const { StartTime: startTime, EndTime: endTime, PageParams: pagination } = data;
    const dateTime = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
    if (typeof startTime !== 'string' || typeof endTime !== 'string'
      || !dateTime.test(startTime) || !dateTime.test(endTime)
      || !Number.isSafeInteger(pagination?.Offset) || pagination.Offset < 0
      || !Number.isSafeInteger(pagination.Limit) || pagination.Limit < 1) return null;

    const impressionMetric = modes[mode].impressionMetric;
    const impressionOrder = [{ Field: impressionMetric, Type: 2 }];
    const originalOrder = data.OrderBy;
    const alreadySorted = Array.isArray(originalOrder) && originalOrder.length === 1
      && originalOrder[0]?.Field === impressionMetric && originalOrder[0]?.Type === 2;
    const withOffset = (offset, period) => JSON.stringify({
      ...data,
      StartTime: period.startTime,
      EndTime: period.endTime,
      PageParams: { ...pagination, Offset: offset, Limit: batchLimit },
      OrderBy: impressionOrder,
    });
    return {
      ...request,
      mode,
      startTime,
      endTime,
      keyForPeriod: (period) => withOffset(0, period),
      replay: (offset, signal, period) => send(
        offset === pagination.Offset && pagination.Limit === batchLimit && alreadySorted
          && startTime === period.startTime && endTime === period.endTime ? body : withOffset(offset, period),
        signal),
    };
  }

  function succeeded(result) {
    return result?.status >= 200 && result.status < 300 && String(result.json?.status_code) === '0';
  }

  function syncAutomaticPeriod() {
    if (!periodAutomatic || batch?.phase === 'running' || batch?.phase === 'complete') return;
    const nextPeriod = recentPageMonthPeriod(templatesForPage());
    if (!samePeriod(selectedPeriod, nextPeriod)) {
      selectedPeriod = nextPeriod;
      batch = undefined;
    }
  }

  function saveTemplate(template, result) {
    if (!succeeded(result) || template.account !== currentPage()) return;
    if (captureAccount !== template.account) {
      templates.clear();
      captureAccount = template.account;
    }
    if ((templates.get(template.mode)?.sequence ?? 0) > template.sequence) return;
    templates.set(template.mode, template);
    if (batch?.account === template.account && !matchesBatch(batch, templatesForPage())) {
      if (batch.phase === 'running') {
        batch.restartRequired = true;
        batch.pauseRequested = true;
        batch.error = '页面筛选已变更，请重新采集。';
      } else {
        batch = undefined;
      }
    }
    syncAutomaticPeriod();
    updatePanel();
  }

  function readXhr(xhr) {
    let json;
    try {
      if (xhr.responseType === 'json') json = xhr.response;
      else if (xhr.responseType === '' || xhr.responseType === 'text') json = JSON.parse(xhr.responseText);
    } catch {
      // Keep the HTTP status when the server returns HTML or an empty body.
    }
    return { status: xhr.status, json };
  }

  async function readFetch(response) {
    let json;
    try {
      json = await response.clone().json();
    } catch {
      // Keep the HTTP status when the body cannot be read as JSON.
    }
    return { status: response.status, json };
  }

  function replayXhr(request, body, withCredentials, signal) {
    return new Promise((resolve, reject) => {
      // The page may wrap XHR after document-start. Keep sending and event
      // registration on the same current chain so queued listeners are flushed.
      const xhr = new window.XMLHttpRequest();
      replayRequests.add(xhr);
      const abort = () => xhr.abort();
      const finish = () => {
        signal.removeEventListener('abort', abort);
        resolve(readXhr(xhr));
      };
      try {
        xhr.open('POST', request.url.href, true);
        xhr.withCredentials = withCredentials;
        for (const [name, value] of request.headers) xhr.setRequestHeader(name, value);
        xhr.addEventListener('loadend', finish, { once: true });
        xhr.addEventListener('load', finish, { once: true });
        signal.addEventListener('abort', abort, { once: true });
        if (signal.aborted) {
          abort();
          finish();
          return;
        }
        xhr.send(body);
      } catch (error) {
        signal.removeEventListener('abort', abort);
        reject(error);
      }
    });
  }

  xhrPrototype.open = function open(method, input) {
    const result = xhrOpen.apply(this, arguments);
    requests.delete(this);
    if (replayRequests.has(this)) return result;
    try {
      const account = currentPage();
      const url = account === null ? null : matchEndpoint(method, input);
      if (url) requests.set(this, { account, url, headers: [] });
    } catch {
      // Leave validation to the page's original transport.
    }
    return result;
  };

  xhrPrototype.setRequestHeader = function setRequestHeader(name, value) {
    const result = xhrSetHeader.apply(this, arguments);
    const request = requests.get(this);
    if (request && !/^cookie2?$/i.test(name)) request.headers.push([String(name), String(value)]);
    return result;
  };

  xhrPrototype.send = function send(body) {
    const request = requests.get(this);
    let listener;
    if (request && request.account === currentPage()) {
      request.sequence = nextSequence++;
      const withCredentials = this.withCredentials;
      const template = readTemplate(request, body,
        (nextBody, signal) => replayXhr(request, nextBody, withCredentials, signal));
      if (template) {
        listener = () => saveTemplate(template, readXhr(this));
        this.addEventListener('loadend', listener, { once: true });
      }
    }

    try {
      return xhrSend.apply(this, arguments);
    } catch (error) {
      if (listener) this.removeEventListener('loadend', listener);
      throw error;
    }
  };

  async function observeFetch(request, snapshot, pending) {
    const body = await snapshot.clone().text();
    const template = readTemplate(request, body, async (nextBody, signal) => {
      const init = { signal };
      if (nextBody !== body) init.body = nextBody;
      return readFetch(await pageFetch.call(window, snapshot.clone(), init));
    });
    if (template) saveTemplate(template, await pending);
  }

  window.fetch = function fetch(input, init) {
    const account = currentPage();
    let snapshot;
    let url;
    try {
      if (account !== null) {
        const isRequest = input instanceof Request;
        url = matchEndpoint(init?.method ?? (isRequest ? input.method : 'GET'), isRequest ? input.url : input);
        if (url) snapshot = new Request(isRequest ? input.clone() : input, init);
      }
    } catch {
      // Leave validation to the page's original transport.
    }

    const pending = pageFetch.apply(this, arguments);
    if (snapshot) {
      const result = pending.then(readFetch, () => ({ status: 0 }));
      observeFetch({ account, url, sequence: nextSequence++ }, snapshot, result).catch(() => {});
    }
    return pending;
  };

  function templatesForPage() {
    return captureAccount === currentPage() ? modeCodes.map((mode) => templates.get(mode)) : [];
  }

  function canResume(run, list) {
    return run?.phase === 'paused' && !run.restartRequired
      && run.startTime === selectedPeriod.startTime && run.endTime === selectedPeriod.endTime
      && matchesBatch(run, list);
  }

  function matchesBatch(run, list) {
    return run.account === currentPage()
      && run.scans.every((scan, index) => scan.queryKey === list[index]?.keyForPeriod(run));
  }

  function planBatch(list) {
    return {
      account: currentPage(),
      ...selectedPeriod,
      phase: 'running',
      scans: list.map((template) => ({
        mode: template.mode,
        queryKey: template.keyForPeriod(selectedPeriod),
        offset: 0,
        total: null,
        pages: 0,
        pageKeys: new Set(),
        matches: new Map(),
        spendCents: 0,
        impressions: 0,
        descending: true,
        lastImpressions: Infinity,
        firstBelowImpressions: null,
        firstBelowImpressionsValue: null,
        complete: false,
      })),
    };
  }

  function readPage(result) {
    if (!succeeded(result)) {
      if (result.timedOut) throw new Error('请求超过 15 秒未完成，请重新打开对应标签页后继续。');
      if (result.status === 429) throw new Error('接口限流，请稍后继续。');
      if (result.status === 401 || result.status === 403) {
        throw new Error(`HTTP ${result.status}，请在对应标签页重新查询后继续。`);
      }
      const detail = result.status ? `HTTP ${result.status}` : '网络请求失败';
      const code = result.json?.status_code;
      throw new Error(code === undefined ? `${detail}，请稍后继续。` : `${detail} 状态码 ${code}，请稍后继续。`);
    }

    const stats = result.json?.data?.StatsData;
    const count = stats?.TotalCount;
    const total = typeof count === 'string' && /^\d+$/.test(count) ? Number(count) : count;
    const rows = total === 0 && stats?.Rows == null ? [] : stats?.Rows;
    if (!Number.isSafeInteger(total) || total < 0 || !Array.isArray(rows)) {
      throw new Error('响应缺少有效的素材总量或列表。');
    }
    return { total, rows };
  }

  function videoLink(row) {
    try {
      const raw = row.Dimensions?.material_content_v2?.Value;
      const link = new URL(JSON.parse(raw)?.video_info?.video_url);
      return link.protocol === 'https:' || link.protocol === 'http:' ? link.href : '';
    } catch {
      return '';
    }
  }

  function checkImpressionOrder(scan, rows, impressionMetric) {
    if (!scan.descending) return;
    for (const [index, row] of rows.entries()) {
      const impressions = row?.Metrics?.[impressionMetric]?.Value;
      if (!Number.isSafeInteger(impressions) || impressions < 0 || impressions > scan.lastImpressions) {
        scan.descending = false;
        scan.firstBelowImpressions = null;
        scan.firstBelowImpressionsValue = null;
        return;
      }
      if (scan.firstBelowImpressions === null && impressions < minImpressions) {
        scan.firstBelowImpressions = scan.offset + index + 1;
        scan.firstBelowImpressionsValue = impressions;
      }
      scan.lastImpressions = impressions;
    }
  }

  function addPage(run, scan, page) {
    const { rows, total } = page;
    if (scan.total !== null && scan.total !== total) {
      run.restartRequired = true;
      throw new Error('分页期间接口总量发生变化，请重新采集。');
    }
    if (scan.offset + rows.length > total || rows.length > batchLimit
      || (!rows.length && scan.offset < total)) {
      throw new Error('接口未返回完整分页，请检查当前查询。');
    }

    const ids = rows.map((row) => String(row?.Dimensions?.material_id?.Value ?? ''));
    if (ids.length && ids.every(Boolean)) {
      const pageKey = JSON.stringify([...ids].sort());
      if (scan.pageKeys.has(pageKey)) throw new Error('接口返回了重复页，已暂停以免漏采。');
      scan.pageKeys.add(pageKey);
    }

    const impressionMetric = modes[scan.mode].impressionMetric;
    checkImpressionOrder(scan, rows, impressionMetric);
    rows.forEach((row, index) => {
      const impressions = row?.Metrics?.[impressionMetric]?.Value;
      if (!Number.isSafeInteger(impressions) || impressions < minImpressions) return;
      const spend = row?.Metrics?.[spendMetric]?.Value;
      if (!Number.isFinite(spend) || spend < minSpend) return;

      const id = ids[index];
      const key = id ? `id:${id}` : `row:${scan.offset + index}`;
      if (scan.matches.has(key)) return;

      const name = row.Dimensions?.material_name_v2?.Value || id || '未命名素材';
      const spendCents = Math.round(spend * 100);
      scan.matches.set(key, { id, name, spendCents, impressions, link: videoLink(row) });
      scan.spendCents += spendCents;
      scan.impressions += impressions;
    });

    scan.total = total;
    scan.offset += rows.length;
    if (rows.length) scan.pages++;
    scan.complete = scan.offset === total || scan.firstBelowImpressions !== null;
  }

  async function replayPage(template, offset, period) {
    const controller = new AbortController();
    let timer;
    const deadline = new Promise((resolve) => {
      timer = setTimeout(() => {
        resolve({ status: 0, timedOut: true });
        controller.abort();
      }, requestTimeoutMs);
    });
    try {
      return await Promise.race([template.replay(offset, controller.signal, period), deadline]);
    } catch {
      return { status: 0 };
    } finally {
      clearTimeout(timer);
    }
  }

  async function collect(run) {
    for (const scan of run.scans) {
      while (!scan.complete && !run.pauseRequested) {
        if (run.restartRequired || !matchesBatch(run, templatesForPage())) {
          throw new Error('页面查询已变更，请恢复原查询或重新采集。');
        }
        const template = templates.get(scan.mode);
        run.activeMode = scan.mode;
        updatePanel();

        const result = await replayPage(template, scan.offset, run);
        if (currentPage() !== run.account) throw new Error('已离开原素材分析页面。');
        if (run.restartRequired || !matchesBatch(run, templatesForPage())) {
          throw new Error('页面筛选已变更，请重新采集。');
        }
        addPage(run, scan, readPage(result));
        updatePanel();

        if (!run.pauseRequested && !run.scans.every((item) => item.complete)) {
          await new Promise((resolve) => setTimeout(resolve, pageDelayMs));
        }
      }
      if (run.pauseRequested) break;
    }
  }

  async function startCollection() {
    const list = templatesForPage();
    if (!list[0] || !list[1] || dateError || batch?.phase === 'running') return;

    const previous = batch?.account === currentPage() ? batch : null;
    const resume = canResume(previous, list);
    if (!resume) {
      videoCollectionWorkbook.begin('qianchuan', currentPage() || 'default', {
        startDate: selectedPeriod.startTime.slice(0, 10),
        endDate: selectedPeriod.endTime.slice(0, 10),
      });
    }
    const run = resume ? previous : planBatch(list);
    batch = run;
    run.phase = 'running';
    run.pauseRequested = false;
    run.error = '';
    run.storageError = '';
    updatePanel();

    try {
      await collect(run);
    } catch (error) {
      run.error = `${modes[run.activeMode]?.label ?? ''} ${error.message}`.trim();
    } finally {
      run.phase = !run.restartRequired && run.scans.every((scan) => scan.complete) ? 'complete' : 'paused';
      run.activeMode = null;
      if (run.phase === 'complete') {
        run.storageError = videoCollectionWorkbook.save('qianchuan', workbookSnapshot(run))
          ? '' : '采集完成，但总表数据保存失败，请刷新页面后重新采集。';
      }
      syncAutomaticPeriod();
      updatePanel();
    }
  }

  function summarize(run) {
    const totals = { scanned: 0, matches: 0, spendCents: 0, impressions: 0 };
    for (const scan of run?.scans ?? []) {
      totals.scanned += scan.offset;
      totals.matches += scan.matches.size;
      totals.spendCents += scan.spendCents;
      totals.impressions += scan.impressions;
    }
    return totals;
  }

  function workbookSnapshot(run) {
    const totals = summarize(run);
    const rows = [];
    for (const scan of run.scans) {
      for (const row of scan.matches.values()) {
        rows.push({
          promotionType: modes[scan.mode].label,
          materialId: row.id,
          title: row.name,
          spendYuan: row.spendCents / 100,
          impressions: row.impressions,
        });
      }
    }
    return {
      scope: run.account || 'default',
      startDate: run.startTime.slice(0, 10),
      endDate: run.endTime.slice(0, 10),
      scanned: totals.scanned,
      rows,
    };
  }

  function csvCell(value) {
    let text = String(value);
    if (/^[\t\r\n]|^\s*[=+@-]/.test(text)) text = `'${text}`;
    return `"${text.replaceAll('"', '""')}"`;
  }

  function csvText(run) {
    const totals = summarize(run);
    const lines = [['推广类型', '素材 ID', '素材名称', '消耗（元）', '展现（次）', '视频链接']];
    for (const scan of run.scans) {
      for (const row of scan.matches.values()) {
        lines.push([modes[scan.mode].label, row.id, row.name, (row.spendCents / 100).toFixed(2), row.impressions, row.link]);
      }
    }
    lines.push(['合计', '', `${totals.matches} 条素材`, (totals.spendCents / 100).toFixed(2), totals.impressions, '']);

    return '\uFEFF' + lines.map((line) => line.map(csvCell).join(',')).join('\r\n');
  }

  function downloadCsv(run) {
    const url = URL.createObjectURL(new Blob([csvText(run)], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `qianchuan-material-${run.startTime.slice(0, 10)}_${run.endTime.slice(0, 10)}.csv`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function scanProgress(mode, template, scan) {
    const label = modes[mode].label;
    if (!scan || scan.total === null) return `${label} ${template ? '已捕获' : '待捕获'}`;
    if (scan.firstBelowImpressions !== null) {
      return `${label} ${scan.pages} 页 展现倒序第 ${numberFormat.format(scan.firstBelowImpressions)} 条 ${numberFormat.format(scan.firstBelowImpressionsValue)} 次，低于 ${numberFormat.format(minImpressions)} 次`;
    }
    const pages = scan.pages + Math.ceil((scan.total - scan.offset) / batchLimit);
    return `${label} ${scan.pages} / ${pages} 页`;
  }

  function panelView() {
    const list = templatesForPage();
    const ready = Boolean(list[0] && list[1]);
    const run = batch?.account === currentPage() ? batch : null;
    const running = run?.phase === 'running';
    const complete = run?.phase === 'complete';
    const differentScopes = list.filter((item) => item && !samePeriod(item, run ?? selectedPeriod));
    const totals = summarize(run);
    const workbookState = videoCollectionWorkbook.summary();
    const rows = [];

    for (const scan of run?.scans ?? []) {
      for (const item of scan.matches.values()) {
        if (rows.length >= previewLimit) break;
        rows.push({
          key: `${scan.mode}:${item.id || rows.length}`,
          type: modes[scan.mode].label,
          name: item.name,
          spend: numberFormat.format(item.spendCents / 100),
          impressions: numberFormat.format(item.impressions),
        });
      }
    }

    let status = '';
    if (dateError) status = dateError;
    else if (running) status = run.pauseRequested ? '正在暂停' : '正在逐页采集，请保持页面打开。';
    else if (complete) status = run.storageError || '两类素材筛选完成，可以下载当前 CSV。';
    else if (run?.phase === 'paused') status = run.error || '采集已暂停，进度已保留。';
    else if (!ready) status = '请分别打开推直播间和推商品，等待捕获列表请求。';

    const scopeNote = differentScopes.length ? [
      ...differentScopes.map((item) => `${modes[item.mode].label} 页面查询 ${periodText(item)}`),
      `采集日期 ${periodText(run ?? selectedPeriod)}`,
      '核对采集序号时，请将对应页面设为相同日期，并按展现从高到低排序。',
    ].join('。') : '';

    return {
      periodLabel: '统计周期',
      periodText: periodText(run ?? selectedPeriod),
      periodAutomatic,
      startDate: selectedPeriod.startTime.slice(0, 10),
      endDate: selectedPeriod.endTime.slice(0, 10),
      running,
      pauseDisabled: Boolean(run?.pauseRequested),
      startDisabled: !ready || Boolean(dateError),
      downloadDisabled: !complete || Boolean(dateError),
      combinedDownloadDisabled: !workbookState.ready,
      workbookItems: workbookState.items,
      workbookStatus: workbookState.message,
      startLabel: run ? (canResume(run, list) ? '继续采集' : '重新采集') : '开始采集',
      scopeNote,
      status,
      progress: modeCodes.map((mode, index) => {
        const scan = run?.scans[index];
        const item = list[index];
        let state = item ? 'ready' : 'waiting';
        if (scan?.complete) state = 'done';
        else if (running && run.activeMode === mode) state = 'active';
        return { key: mode, state, text: scanProgress(mode, item, scan) };
      }),
      filters: [
        `消耗 ≥ ${numberFormat.format(minSpend)} 元`,
        `展现 ≥ ${numberFormat.format(minImpressions)} 次`,
      ],
      kpis: [
        { label: '已采集', value: numberFormat.format(totals.scanned), unit: '条' },
        { label: '达标素材', value: numberFormat.format(totals.matches), unit: '条', highlight: true },
        { label: '消耗', value: numberFormat.format(totals.spendCents / 100), unit: '元' },
        { label: '展现', value: numberFormat.format(totals.impressions), unit: '次' },
      ],
      progressPercent: complete ? 100 : (run?.scans?.length
        ? Math.min(95, Math.round((run.scans.reduce((sum, s) => sum + s.pages, 0) / Math.max(1, run.scans.reduce((sum, s) => sum + (s.total ? Math.ceil(s.total / batchLimit) : 1), 0))) * 100))
        : 0),
      countText: [
        `已采集 ${numberFormat.format(totals.scanned)} 条`,
        `达标 ${numberFormat.format(totals.matches)} 条`,
        `消耗 ${numberFormat.format(totals.spendCents / 100)} 元`,
        `展现 ${numberFormat.format(totals.impressions)} 次`,
        totals.matches > previewLimit ? `预览前 ${previewLimit} 条` : '',
      ].filter(Boolean).join(' '),
      rows,
      columns: [
        { prop: 'type', label: '推广类型', width: 92 },
        { prop: 'name', label: '素材', minWidth: 180 },
        { prop: 'spend', label: '消耗（元）', width: 104, align: 'right', numeric: true },
        { prop: 'impressions', label: '展现（次）', width: 110, align: 'right', numeric: true },
      ],
    };
  }

  function applyPanelDates(startDate, endDate, automatic = false) {
    if (batch?.phase === 'running') return;
    periodAutomatic = automatic === true;
    dateError = '';
    if (!isDate(startDate) || !isDate(endDate) || startDate > endDate) {
      dateError = '请选择有效的起止日期，结束日期须等于或晚于开始日期。';
    } else {
      const nextPeriod = {
        startTime: `${startDate} 00:00:00`,
        endTime: `${endDate} 23:59:59`,
      };
      if (!samePeriod(selectedPeriod, nextPeriod)) {
        selectedPeriod = nextPeriod;
        batch = undefined;
      }
    }
    updatePanel();
  }

  function mountPanel() {
    const controller = videoCollectionLiquidUi.mount({
      hostId: 'qianchuan-material-collector',
      config: {
        title: '素材采集',
        launcherLabel: '素材采集',
        panelTop: '250px',
        downloadLabel: '下载当前 CSV',
        tableLabel: '符合筛选条件的素材',
      },
      initialPeriod: () => ({
        startDate: selectedPeriod.startTime.slice(0, 10),
        endDate: selectedPeriod.endTime.slice(0, 10),
      }),
      getView: panelView,
      isVisible: () => currentPage() !== null,
      applyDates: applyPanelDates,
      monthPeriod: () => {
        const current = recentPageMonthPeriod(templatesForPage());
        return {
          startDate: current.startTime.slice(0, 10),
          endDate: current.endTime.slice(0, 10),
        };
      },
      start: startCollection,
      pause: () => {
        if (batch?.phase !== 'running') return;
        batch.pauseRequested = true;
        updatePanel();
      },
      download: () => {
        if (batch?.account === currentPage() && batch.phase === 'complete') downloadCsv(batch);
      },
      downloadCombined: () => videoCollectionWorkbook.downloadCombined(),
      subscribe: (callback) => videoCollectionWorkbook.subscribe(callback),
      unsubscribe: (listeners) => videoCollectionWorkbook.unsubscribe(listeners),
    });
    updatePanel = controller.refresh;
    controller.refresh();
    return controller.open;
  }
  let openPanel;
  if (document.body) openPanel = mountPanel();
  else document.addEventListener('DOMContentLoaded', () => { openPanel = mountPanel(); }, { once: true });
  GM_registerMenuCommand('打开采集面板', () => openPanel?.());
})();

(() => {
  'use strict';

  const origin = 'https://compass.jinritemai.com';
  const pagePath = '/shop/video/self';
  const listPath = '/compass_api/shop/video/overview/video_list';
  const pageWindow = typeof unsafeWindow === 'object' && unsafeWindow ? unsafeWindow : window;
  const minWatchCount = 20000;
  const maxPages = 500;
  const pageDelayMs = 300;
  const requestTimeoutMs = 15000;
  const previewLimit = 50;
  const numberFormat = new Intl.NumberFormat('zh-CN');

  function onSelfVideoPage() {
    return location.origin === origin && location.pathname === pagePath;
  }

  if (!onSelfVideoPage()) return;

  function isDate(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(value) && !value.startsWith('0000')
      && new Date(`${value}T00:00:00Z`).toJSON()?.slice(0, 10) === value;
  }

  function currentMonthPeriod(now = new Date()) {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(now);
    const date = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    const yesterday = new Date(Date.UTC(Number(date.year), Number(date.month) - 1, Number(date.day) - 1));
    const endDate = yesterday.toISOString().slice(0, 10);
    return { startDate: `${endDate.slice(0, 7)}-01`, endDate };
  }

  function recentPageMonthPeriod(source) {
    const current = currentMonthPeriod();
    if (!source || source.endDate.slice(0, 7) !== current.endDate.slice(0, 7)) return current;
    const lag = (Date.parse(`${current.endDate}T00:00:00Z`)
      - Date.parse(`${source.endDate}T00:00:00Z`)) / 86400000;
    return Number.isInteger(lag) && lag >= 0 && lag <= 7
      ? { startDate: current.startDate, endDate: source.endDate }
      : current;
  }

  function samePeriod(left, right) {
    return left?.startDate === right?.startDate && left?.endDate === right?.endDate;
  }

  function periodText(period) {
    return `${period.startDate} 至 ${period.endDate}`;
  }

  function apiDate(date) {
    return `${date.replaceAll('-', '/')} 00:00:00`;
  }

  function dateFromApi(value) {
    const match = /^(\d{4})[/-](\d{2})[/-](\d{2}) 00:00:00$/.exec(String(value));
    if (!match) return '';
    const date = `${match[1]}-${match[2]}-${match[3]}`;
    return isDate(date) ? date : '';
  }

  function positiveInteger(value) {
    const number = typeof value === 'string' && /^\d+$/.test(value) ? Number(value) : value;
    return Number.isSafeInteger(number) && number > 0 ? number : null;
  }

  function nonNegativeInteger(value) {
    const number = typeof value === 'string' && /^\d+$/.test(value) ? Number(value) : value;
    return Number.isSafeInteger(number) && number >= 0 ? number : null;
  }

  let template;
  let nextSequence = 1;
  let batch;
  let selectedPeriod = currentMonthPeriod();
  let periodAutomatic = true;
  let dateError = '';
  let updatePanel = () => {};
  const requests = new WeakMap();
  const replayRequests = new WeakSet();

  function matchEndpoint(method, input) {
    if (String(method).toUpperCase() !== 'GET') return null;
    const url = new URL(input, location.href);
    return url.origin === origin && url.pathname === listPath ? url : null;
  }

  function replayUrl(source, pageNo, period, sourcePeriod) {
    const url = new URL(source.href);
    url.searchParams.set('begin_date', apiDate(period.startDate));
    url.searchParams.set('end_date', apiDate(period.endDate));
    if (!samePeriod(period, sourcePeriod)) url.searchParams.set('date_type', '4');
    url.searchParams.set('page_no', String(pageNo));
    url.searchParams.set('sort_field', 'watch_cnt');
    url.searchParams.set('is_asc', 'false');
    return url;
  }

  function readTemplate(request, body, withCredentials) {
    const pageNo = positiveInteger(request.url.searchParams.get('page_no'));
    const pageSize = positiveInteger(request.url.searchParams.get('page_size'));
    const startDate = dateFromApi(request.url.searchParams.get('begin_date'));
    const endDate = dateFromApi(request.url.searchParams.get('end_date'));
    if (!pageNo || !pageSize || !startDate || !endDate || startDate > endDate) return null;

    const sourcePeriod = { startDate, endDate };
    const source = new URL(request.url.href);
    return {
      sequence: request.sequence,
      source,
      sourcePeriod,
      pageSize,
      headers: request.headers.map(([name, value]) => [name, value]),
      body,
      withCredentials,
      keyForPeriod: (period) => replayUrl(source, 1, period, sourcePeriod).href,
      replay: (nextPage, signal, period) => replayXhr({
        source, sourcePeriod, headers: request.headers, body, withCredentials,
      }, nextPage, period, signal),
    };
  }

  function readXhr(xhr) {
    let json;
    try {
      if (xhr.responseType === 'json') json = xhr.response;
      else if (xhr.responseType === '' || xhr.responseType === 'text') json = JSON.parse(xhr.responseText);
    } catch {
      // Keep the HTTP status when the response is not JSON.
    }
    return { status: xhr.status, json };
  }

  function succeeded(result) {
    return result?.status >= 200 && result.status < 300 && String(result.json?.st) === '0';
  }

  function tableFrom(result) {
    return result?.json?.data?.module_data?.core_data_0?.compass_general_table_value;
  }

  function replayXhr(source, pageNo, period, signal) {
    return new Promise((resolve, reject) => {
      const xhr = new pageWindow.XMLHttpRequest();
      replayRequests.add(xhr);
      let settled = false;
      const abort = () => xhr.abort();
      const finish = () => {
        if (settled) return;
        settled = true;
        signal.removeEventListener('abort', abort);
        resolve(readXhr(xhr));
      };
      try {
        const url = replayUrl(source.source, pageNo, period, source.sourcePeriod);
        xhr.open('GET', url.href, true);
        xhr.withCredentials = source.withCredentials;
        for (const [name, value] of source.headers) xhr.setRequestHeader(name, value);
        xhr.addEventListener('loadend', finish, { once: true });
        signal.addEventListener('abort', abort, { once: true });
        if (signal.aborted) {
          abort();
          finish();
          return;
        }
        xhr.send(source.body);
      } catch (error) {
        signal.removeEventListener('abort', abort);
        reject(error);
      }
    });
  }

  function matchesBatch(run) {
    return onSelfVideoPage() && template && run.queryKey === template.keyForPeriod(run);
  }

  function collectionScope() {
    const accountType = template?.source.searchParams.get('account_type') || 'default';
    const authorId = template?.source.searchParams.get('author_id') || 'default';
    return `${accountType}:${authorId}`;
  }

  function syncAutomaticPeriod() {
    if (!periodAutomatic || batch?.phase === 'running' || batch?.phase === 'complete') return;
    const nextPeriod = recentPageMonthPeriod(template?.sourcePeriod);
    if (!samePeriod(selectedPeriod, nextPeriod)) {
      selectedPeriod = nextPeriod;
      batch = undefined;
    }
  }

  function saveTemplate(candidate, result) {
    if (!succeeded(result) || !Array.isArray(tableFrom(result)?.data) || !onSelfVideoPage()) return;
    if ((template?.sequence ?? 0) > candidate.sequence) return;
    template = candidate;
    if (batch && !matchesBatch(batch)) {
      if (batch.phase === 'running') {
        batch.restartRequired = true;
        batch.pauseRequested = true;
        batch.error = '页面查询已变更，请重新采集。';
      } else {
        batch = undefined;
      }
    }
    syncAutomaticPeriod();
    updatePanel();
  }

  function observeOpen(xhr, open, args) {
    const result = open.apply(xhr, args);
    requests.delete(xhr);
    try {
      const url = onSelfVideoPage() ? matchEndpoint(args[0], args[1]) : null;
      if (url) requests.set(xhr, { url, headers: [] });
    } catch {
      // Leave invalid requests to the page's original transport.
    }
    return result;
  }

  function observeSetHeader(xhr, setRequestHeader, args) {
    const result = setRequestHeader.apply(xhr, args);
    const request = requests.get(xhr);
    const [name, value] = args;
    if (request && !/^cookie2?$/i.test(name)) request.headers.push([String(name), String(value)]);
    return result;
  }

  function observeSend(xhr, send, args) {
    const request = requests.get(xhr);
    requests.delete(xhr);
    let listener;
    if (request && onSelfVideoPage()) {
      request.sequence = nextSequence++;
      const candidate = readTemplate(request, args[0], xhr.withCredentials);
      if (candidate) {
        listener = () => saveTemplate(candidate, readXhr(xhr));
        xhr.addEventListener('loadend', listener, { once: true });
      }
    }
    try {
      return send.apply(xhr, args);
    } catch (error) {
      if (listener) xhr.removeEventListener('loadend', listener);
      throw error;
    }
  }

  const xhrObservers = [
    ['open', observeOpen, new WeakSet()],
    ['setRequestHeader', observeSetHeader, new WeakSet()],
    ['send', observeSend, new WeakSet()],
  ];
  const captureHooks = new WeakSet();

  function installCapture() {
    if (!onSelfVideoPage()) return;
    try {
      const prototype = pageWindow.XMLHttpRequest?.prototype;
      if (!prototype) return;
      for (const [method, observe, active] of xhrObservers) {
        const current = prototype[method];
        if (typeof current !== 'function' || captureHooks.has(current)) continue;
        const hook = function (...args) {
          if (replayRequests.has(this) || active.has(this)) return current.apply(this, args);
          active.add(this);
          try {
            return observe(this, current, args);
          } finally {
            active.delete(this);
          }
        };
        captureHooks.add(hook);
        prototype[method] = hook;
      }
    } catch {
      // The page can initialize its transport later. Try again at the next interaction.
    }
  }

  installCapture();
  document.addEventListener('DOMContentLoaded', installCapture, { once: true });
  // The active XHR methods may change while the page initializes. Recheck before native queries.
  for (const event of ['pointerdown', 'keydown', 'click']) {
    document.addEventListener(event, installCapture, true);
  }

  function readPage(result) {
    if (!succeeded(result)) {
      if (result.timedOut) throw new Error('请求超过 15 秒未完成，请重新查询列表后继续。');
      if (result.status === 429) throw new Error('接口限流，请稍后继续。');
      if (result.status === 401 || result.status === 403) {
        throw new Error(`HTTP ${result.status}，请重新查询列表后继续。`);
      }
      const detail = result.status ? `HTTP ${result.status}` : '网络请求失败';
      const code = result.json?.st ?? result.json?.code;
      throw new Error(code === undefined ? `${detail}，请稍后继续。` : `${detail} 状态码 ${code}，请稍后继续。`);
    }

    const table = tableFrom(result);
    const rows = table?.data;
    const page = table?.page_result;
    const pageNo = positiveInteger(page?.page_no);
    const pageSize = positiveInteger(page?.page_size);
    const total = nonNegativeInteger(page?.total);
    if (!Array.isArray(rows) || !pageNo || !pageSize || total === null) {
      throw new Error('响应缺少有效的视频列表或分页信息。');
    }
    return { rows, pageNo, pageSize, total };
  }

  function rowData(row) {
    const video = row?.cell_info?.video?.video;
    const watchCount = nonNegativeInteger(row?.cell_info?.watch_cnt?.index_values?.value?.value);
    return {
      id: video?.video_id == null ? '' : String(video.video_id),
      title: typeof video?.video_title === 'string' && video.video_title.trim()
        ? video.video_title.trim() : '未命名视频',
      watchCount,
    };
  }

  function planBatch() {
    return {
      ...selectedPeriod,
      phase: 'running',
      queryKey: template.keyForPeriod(selectedPeriod),
      pageNo: 1,
      pageSize: template.pageSize,
      total: null,
      pages: 0,
      scanned: 0,
      pageKeys: new Set(),
      matches: new Map(),
      descending: true,
      lastWatchCount: Infinity,
      firstOutside: null,
      complete: false,
    };
  }

  function canResume(run) {
    return run?.phase === 'paused' && !run.restartRequired
      && samePeriod(run, selectedPeriod) && matchesBatch(run);
  }

  function addPage(run, page) {
    if (page.pageNo !== run.pageNo || page.pageSize !== run.pageSize) {
      throw new Error('接口返回的页码或每页条数与请求不一致。');
    }
    if (run.total !== null && page.total !== run.total) {
      run.restartRequired = true;
      throw new Error('分页期间接口总量发生变化，请重新采集。');
    }
    if (run.scanned + page.rows.length > page.total || page.rows.length > page.pageSize
      || (!page.rows.length && run.scanned < page.total)) {
      throw new Error('接口未返回完整分页，请检查当前查询。');
    }

    const values = page.rows.map(rowData);
    const ids = values.map((item) => item.id);
    if (ids.length && ids.every(Boolean)) {
      const pageKey = JSON.stringify([...ids].sort());
      if (run.pageKeys.has(pageKey)) throw new Error('接口返回了重复页，已暂停以免漏采。');
      run.pageKeys.add(pageKey);
    }

    for (const [index, item] of values.entries()) {
      if (item.watchCount === null || item.watchCount > run.lastWatchCount) {
        run.descending = false;
        run.firstOutside = null;
      } else if (run.descending && run.firstOutside === null && item.watchCount <= minWatchCount) {
        run.firstOutside = {
          position: run.scanned + index + 1,
          watchCount: item.watchCount,
        };
      }
      if (item.watchCount !== null) run.lastWatchCount = item.watchCount;
      if (item.watchCount === null || item.watchCount <= minWatchCount) continue;

      const key = item.id ? `id:${item.id}` : `row:${run.pageNo}:${index}`;
      if (!run.matches.has(key)) {
        run.matches.set(key, { title: item.title, watchCount: item.watchCount });
      }
    }

    run.total = page.total;
    run.scanned += page.rows.length;
    if (page.rows.length) run.pages++;
    run.pageNo++;
    run.complete = run.scanned === page.total || (run.descending && run.firstOutside !== null);
  }

  async function replayPage(run) {
    const controller = new AbortController();
    let timer;
    const deadline = new Promise((resolve) => {
      timer = setTimeout(() => {
        resolve({ status: 0, timedOut: true });
        controller.abort();
      }, requestTimeoutMs);
    });
    try {
      return await Promise.race([template.replay(run.pageNo, controller.signal, run), deadline]);
    } catch {
      return { status: 0 };
    } finally {
      clearTimeout(timer);
    }
  }

  async function collect(run) {
    while (!run.complete && !run.pauseRequested) {
      if (run.pages >= maxPages) throw new Error(`已达到 ${maxPages} 页上限，请缩小日期范围。`);
      if (run.restartRequired || !matchesBatch(run)) {
        throw new Error('页面查询已变更，请恢复原查询或重新采集。');
      }
      updatePanel();
      const result = await replayPage(run);
      if (!onSelfVideoPage()) throw new Error('已离开自营视频明细页面。');
      if (run.restartRequired || !matchesBatch(run)) {
        throw new Error('页面查询已变更，请重新采集。');
      }
      addPage(run, readPage(result));
      updatePanel();
      if (!run.complete && !run.pauseRequested) {
        await new Promise((resolve) => setTimeout(resolve, pageDelayMs));
      }
    }
  }

  async function startCollection() {
    if (!template || dateError || batch?.phase === 'running') return;
    const resume = canResume(batch);
    if (!resume) videoCollectionWorkbook.begin('compass', collectionScope(), selectedPeriod);
    const run = resume ? batch : planBatch();
    batch = run;
    run.phase = 'running';
    run.pauseRequested = false;
    run.error = '';
    run.storageError = '';
    updatePanel();
    try {
      await collect(run);
    } catch (error) {
      run.error = error.message;
    } finally {
      run.phase = !run.restartRequired && run.complete ? 'complete' : 'paused';
      if (run.phase === 'complete') {
        run.storageError = videoCollectionWorkbook.save('compass', workbookSnapshot(run))
          ? '' : '采集完成，但总表数据保存失败，请刷新页面后重新采集。';
      }
      syncAutomaticPeriod();
      updatePanel();
    }
  }

  function csvCell(value) {
    let text = String(value);
    if (/^[\t\r\n]|^\s*[=+@-]/.test(text)) text = `'${text}`;
    return `"${text.replaceAll('"', '""')}"`;
  }

  function sortedMatches(run) {
    return [...(run?.matches.values() ?? [])]
      .sort((left, right) => right.watchCount - left.watchCount || left.title.localeCompare(right.title, 'zh-CN'));
  }

  function workbookSnapshot(run) {
    return {
      scope: collectionScope(),
      startDate: run.startDate,
      endDate: run.endDate,
      scanned: run.scanned,
      rows: sortedMatches(run).map((row) => ({
        title: row.title,
        watchCount: row.watchCount,
      })),
    };
  }

  function csvText(run) {
    const lines = [['标题', '观看次数']];
    for (const row of sortedMatches(run)) lines.push([row.title, row.watchCount]);
    return '\uFEFF' + lines.map((line) => line.map(csvCell).join(',')).join('\r\n');
  }

  function downloadCsv(run) {
    const url = URL.createObjectURL(new Blob([csvText(run)], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `compass-self-video-${run.startDate}_${run.endDate}.csv`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function progressText(run) {
    if (!template) return '自营视频 待捕获';
    if (!run || run.total === null) return '自营视频 已捕获';
    if (run.firstOutside !== null) {
      return `自营视频 ${run.pages} 页 第 ${numberFormat.format(run.firstOutside.position)} 条观看次数 ${numberFormat.format(run.firstOutside.watchCount)} 次，不大于 ${numberFormat.format(minWatchCount)} 次`;
    }
    const totalPages = Math.ceil(run.total / run.pageSize);
    return `自营视频 ${run.pages} / ${numberFormat.format(totalPages)} 页`;
  }

  function panelView() {
    const run = batch;
    const running = run?.phase === 'running';
    const complete = run?.phase === 'complete';
    const differentScope = template && !samePeriod(template.sourcePeriod, run ?? selectedPeriod);
    const matches = sortedMatches(run);
    const workbookState = videoCollectionWorkbook.summary();
    let status = '';

    if (dateError) status = dateError;
    else if (running) status = run.pauseRequested ? '正在暂停' : '正在逐页采集，请保持页面打开。';
    else if (complete) status = run.storageError || '自营视频筛选完成，可以下载当前 CSV。';
    else if (run?.phase === 'paused') status = run.error || '采集已暂停，进度已保留。';
    else if (!template) status = '请刷新页面或重新查询视频明细，等待捕获列表请求。';

    return {
      periodLabel: '采集日期',
      periodText: periodText(run ?? selectedPeriod),
      periodAutomatic,
      startDate: selectedPeriod.startDate,
      endDate: selectedPeriod.endDate,
      running,
      pauseDisabled: Boolean(run?.pauseRequested),
      startDisabled: !template || Boolean(dateError),
      downloadDisabled: !complete || Boolean(dateError),
      combinedDownloadDisabled: !workbookState.ready,
      workbookItems: workbookState.items,
      workbookStatus: workbookState.message,
      startLabel: run ? (canResume(run) ? '继续采集' : '重新采集') : '开始采集',
      scopeNote: differentScope
        ? `最近页面查询 ${periodText(template.sourcePeriod)}。采集日期 ${periodText(run ?? selectedPeriod)}。核对序号时，请将页面设为相同日期并按观看次数从高到低排序。`
        : '',
      status,
      progress: [{
        key: 'self-video',
        state: !template ? 'waiting' : complete ? 'done' : running ? 'active' : 'ready',
        text: progressText(run),
      }],
      filters: [`观看次数 > ${numberFormat.format(minWatchCount)} 次`],
      kpis: [
        { label: '已采集', value: numberFormat.format(run?.scanned ?? 0), unit: '条' },
        { label: '达标视频', value: numberFormat.format(run?.matches.size ?? 0), unit: '条', highlight: true },
      ],
      progressPercent: complete ? 100 : (run?.total
        ? Math.min(95, Math.round((run.pages / Math.max(1, Math.ceil(run.total / run.pageSize))) * 100))
        : 0),
      countText: [
        `已采集 ${numberFormat.format(run?.scanned ?? 0)} 条`,
        `达标 ${numberFormat.format(run?.matches.size ?? 0)} 条`,
        matches.length > previewLimit ? `预览前 ${previewLimit} 条` : '',
      ].filter(Boolean).join(' '),
      rows: matches.slice(0, previewLimit).map((item, index) => ({
        key: `${index}:${item.title}:${item.watchCount}`,
        title: item.title,
        watchCount: numberFormat.format(item.watchCount),
      })),
      columns: [
        { prop: 'title', label: '标题', minWidth: 260 },
        { prop: 'watchCount', label: '观看次数', width: 120, align: 'right', numeric: true },
      ],
    };
  }

  function applyPanelDates(startDate, endDate, automatic = false) {
    if (batch?.phase === 'running') return;
    periodAutomatic = automatic === true;
    dateError = '';
    if (!isDate(startDate) || !isDate(endDate) || startDate > endDate) {
      dateError = '请选择有效的起止日期，结束日期须等于或晚于开始日期。';
    } else {
      const nextPeriod = { startDate, endDate };
      if (!samePeriod(selectedPeriod, nextPeriod)) {
        selectedPeriod = nextPeriod;
        batch = undefined;
      }
    }
    updatePanel();
  }

  function mountPanel() {
    const controller = videoCollectionLiquidUi.mount({
      hostId: 'compass-self-video-collector',
      config: {
        title: '自营视频采集',
        launcherLabel: '自营视频采集',
        panelTop: '300px',
        downloadLabel: '下载当前 CSV',
        tableLabel: '符合筛选条件的自营视频',
      },
      initialPeriod: () => ({
        startDate: selectedPeriod.startDate,
        endDate: selectedPeriod.endDate,
      }),
      getView: panelView,
      isVisible: onSelfVideoPage,
      applyDates: applyPanelDates,
      monthPeriod: () => recentPageMonthPeriod(template?.sourcePeriod),
      start: startCollection,
      pause: () => {
        if (batch?.phase !== 'running') return;
        batch.pauseRequested = true;
        updatePanel();
      },
      download: () => {
        if (batch?.phase === 'complete') downloadCsv(batch);
      },
      downloadCombined: () => videoCollectionWorkbook.downloadCombined(),
      subscribe: (callback) => videoCollectionWorkbook.subscribe(callback),
      unsubscribe: (listeners) => videoCollectionWorkbook.unsubscribe(listeners),
    });
    updatePanel = controller.refresh;
    controller.refresh();
    return controller.open;
  }
  let openPanel;
  if (document.body) openPanel = mountPanel();
  else document.addEventListener('DOMContentLoaded', () => { openPanel = mountPanel(); }, { once: true });
  GM_registerMenuCommand('打开自营视频采集面板', () => openPanel?.());
})();
