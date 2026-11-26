// ==UserScript==
// @name         千川素材请求采集
// @namespace    VideoCollection
// @version      1.4.0
// @description  采集千川素材与抖店自营视频，按页面指标排序筛选后导出 CSV。
//
// @match        https://qianchuan.jinritemai.com/dataV3/roi2-material-analysis*
// @match        https://compass.jinritemai.com/shop/video/self*
// @run-at       document-start
// @sandbox      raw
// @noframes
// @grant        GM_registerMenuCommand
// ==/UserScript==

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
    if (!periodAutomatic || batch?.phase === 'running') return;
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
    const run = canResume(previous, list) ? previous : planBatch(list);
    batch = run;
    run.phase = 'running';
    run.pauseRequested = false;
    run.error = '';
    updatePanel();

    try {
      await collect(run);
    } catch (error) {
      run.error = `${modes[run.activeMode]?.label ?? ''} ${error.message}`.trim();
    } finally {
      run.phase = !run.restartRequired && run.scans.every((scan) => scan.complete) ? 'complete' : 'paused';
      run.activeMode = null;
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

  function csvCell(value) {
    let text = String(value);
    if (/^[\t\r\n]|^\s*[=+@-]/.test(text)) text = `'${text}`;
    return `"${text.replaceAll('"', '""')}"`;
  }

  function downloadCsv(run) {
    const totals = summarize(run);
    const lines = [['推广类型', '素材 ID', '素材名称', '消耗（元）', '展现（次）', '视频链接']];
    for (const scan of run.scans) {
      for (const row of scan.matches.values()) {
        lines.push([modes[scan.mode].label, row.id, row.name, (row.spendCents / 100).toFixed(2), row.impressions, row.link]);
      }
    }
    lines.push(['合计', '', `${totals.matches} 条素材`, (totals.spendCents / 100).toFixed(2), totals.impressions, '']);

    const csv = '\uFEFF' + lines.map((line) => line.map(csvCell).join(',')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
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

  function panelMarkup() {
    return `
      <style>
        :host {
          --collector-ink: #1d293a;
          --collector-muted: #516176;
          --collector-border: #dbe2eb;
          --collector-blue: #1755c7;
          --collector-tint: #f4f7fb;
          font-family: inherit;
          font-size: 13px;
          line-height: 1.45;
        }

        *, *::before, *::after { box-sizing: border-box; }
        [hidden] { display: none !important; }

        .collector-launcher,
        .collector-panel {
          position: fixed;
          z-index: 2147483600;
          inset-inline-end: 18px;
          inset-block-end: 18px;
          color: var(--collector-ink);
        }

        button, input { font: inherit; }
        button { cursor: pointer; }

        button {
          min-height: 34px;
          padding: 5px 12px;
          color: var(--collector-ink);
          background: #fff;
          border: 1px solid var(--collector-border);
          border-radius: 6px;
        }

        button:hover:not(:disabled) { background: var(--collector-tint); }
        button:disabled { color: #687689; opacity: 0.7; cursor: default; }

        button:focus-visible, input:focus-visible {
          outline: 2px solid var(--collector-blue);
          outline-offset: 2px;
        }

        .collector-primary {
          color: #fff;
          background: var(--collector-blue);
          border-color: var(--collector-blue);
          font-weight: 600;
        }

        .collector-primary:hover:not(:disabled) {
          color: #fff;
          background: #1248ad;
        }

        .collector-launcher {
          min-height: 40px;
          padding-inline: 16px;
          box-shadow: 0 3px 12px rgb(28 49 83 / 20%);
        }

        .collector-panel {
          width: min(480px, calc(100vw - 36px));
          max-height: min(620px, calc(100vh - 36px));
          max-height: min(620px, calc(100dvh - 36px));
          overflow: auto;
          overscroll-behavior: contain;
          background: #fff;
          border: 1px solid var(--collector-border);
          border-radius: 10px;
          box-shadow: 0 10px 30px rgb(24 42 68 / 18%);
        }

        .collector-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          padding: 12px 16px;
          border-bottom: 1px solid var(--collector-border);
        }

        .collector-header h2 { margin: 0; font-size: 16px; font-weight: 650; }
        .collector-close { min-height: 30px; padding: 3px 9px; }

        .collector-body {
          display: grid;
          gap: 12px;
          padding: 14px 16px 16px;
        }

        .collector-period { margin: 0; color: var(--collector-muted); overflow-wrap: anywhere; }

        .collector-period-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
        }

        .collector-date-fields {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 10px;
        }

        .collector-date-fields label {
          display: grid;
          gap: 6px;
          min-width: 0;
        }

        .collector-date-fields input {
          width: 100%;
          min-width: 0;
          height: 34px;
          padding: 4px 8px;
          color: var(--collector-ink);
          background: #fff;
          border: 1px solid var(--collector-border);
          border-radius: 6px;
        }

        .collector-date-fields input:disabled {
          color: var(--collector-muted);
          background: var(--collector-tint);
        }

        .collector-progress {
          display: grid;
          gap: 6px;
          margin: 0;
          padding: 0;
          list-style: none;
        }

        .collector-progress li {
          min-height: 32px;
          padding: 6px 8px;
          background: var(--collector-tint);
          border-radius: 5px;
          font-variant-numeric: tabular-nums;
          overflow-wrap: anywhere;
        }

        .collector-actions, .collector-filters {
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          gap: 8px;
        }

        .collector-details {
          display: grid;
          gap: 8px;
          padding-top: 12px;
          border-top: 1px solid var(--collector-border);
        }

        .collector-filters, .collector-status, .collector-count { margin: 0; }
        .collector-filters { column-gap: 12px; row-gap: 4px; }
        .collector-filters, .collector-status { color: var(--collector-muted); }
        .collector-status { overflow-wrap: anywhere; }
        .collector-status:empty { display: none; }
        .collector-scope-note {
          padding: 8px 10px;
          color: #7a4f00;
          background: #fff8e6;
          border-inline-start: 3px solid #e5a400;
          border-radius: 5px;
        }
        .collector-count { color: var(--collector-ink); font-weight: 600; }
        .collector-filters span { white-space: nowrap; }

        .collector-table-scroll {
          max-height: 230px;
          overflow: auto;
          overscroll-behavior: contain;
          border: 1px solid var(--collector-border);
          border-radius: 6px;
        }

        table {
          width: 100%;
          table-layout: fixed;
          border-spacing: 0;
          text-align: left;
        }

        th, td {
          padding: 7px 8px;
          vertical-align: top;
          overflow-wrap: anywhere;
          border-bottom: 1px solid var(--collector-border);
        }

        th {
          position: sticky;
          inset-block-start: 0;
          background: var(--collector-tint);
          color: var(--collector-muted);
          font-weight: 650;
          white-space: nowrap;
        }

        th:first-child { width: 21%; }

        th:nth-child(n + 3), td:nth-child(n + 3) {
          width: 22%;
          text-align: right;
          font-variant-numeric: tabular-nums;
        }

        tbody tr:last-child td { border-bottom: 0; }

        @media (max-width: 400px) {
          .collector-launcher, .collector-panel { inset-inline-end: 12px; inset-block-end: 12px; }
          .collector-panel { width: calc(100vw - 24px); max-height: calc(100dvh - 24px); }
          .collector-header, .collector-body { padding-inline: 12px; }
          th, td { padding-inline: 5px; }
          th { white-space: normal; }
          th:first-child { width: 22%; }
        }
      </style>
      <button id="collector-open" class="collector-launcher collector-primary" type="button" aria-controls="collector-panel" aria-expanded="false">素材采集</button>
      <section id="collector-panel" class="collector-panel" aria-labelledby="collector-title" hidden>
        <header class="collector-header">
          <h2 id="collector-title">素材采集</h2>
          <button id="collector-close" class="collector-close" type="button" aria-label="关闭素材采集面板">关闭</button>
        </header>
        <div class="collector-body">
          <div class="collector-period-row">
            <p id="collector-period" class="collector-period">采集日期</p>
            <button id="collector-month" type="button" title="本月 1 日至最近页面查询截止日，无近期查询时截至昨天。每月 1 日选取上月。">本月</button>
          </div>
          <div class="collector-date-fields">
            <label>开始日期<input id="collector-date-start" type="date" required></label>
            <label>结束日期<input id="collector-date-end" type="date" required></label>
          </div>
          <p id="collector-scope-note" class="collector-status collector-scope-note" role="status" hidden></p>
          <ul class="collector-progress">
            <li id="collector-live">推直播间 待捕获</li>
            <li id="collector-product">推商品 待捕获</li>
          </ul>
          <div class="collector-actions">
            <button id="collector-start" class="collector-primary" type="button" disabled>开始采集</button>
            <button id="collector-stop" type="button" hidden>暂停采集</button>
            <button id="collector-download" type="button" disabled>下载 CSV</button>
          </div>
          <p id="collector-status" class="collector-status" role="status"></p>
          <div class="collector-details">
            <p class="collector-filters"><span>消耗 ≥ ${numberFormat.format(minSpend)} 元</span><span>展现 ≥ ${numberFormat.format(minImpressions)} 次</span></p>
            <p id="collector-count" class="collector-count">已采集 0 条 达标 0 条</p>
            <div class="collector-table-scroll" hidden>
              <table aria-label="符合筛选条件的素材">
                <thead>
                  <tr><th scope="col">推广类型</th><th scope="col">素材</th><th scope="col">消耗（元）</th><th scope="col">展现（次）</th></tr>
                </thead>
                <tbody id="collector-rows"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
    `;
  }

  function mountPanel() {
    const host = document.createElement('div');
    host.id = 'qianchuan-material-collector';
    const shadow = host.attachShadow({ mode: 'open' });
    shadow.innerHTML = panelMarkup();
    const openButton = shadow.getElementById('collector-open');
    const panel = shadow.getElementById('collector-panel');
    const closeButton = shadow.getElementById('collector-close');
    const dateStart = shadow.getElementById('collector-date-start');
    const dateEnd = shadow.getElementById('collector-date-end');
    const month = shadow.getElementById('collector-month');
    const scopeNote = shadow.getElementById('collector-scope-note');
    dateStart.value = selectedPeriod.startTime.slice(0, 10);
    dateEnd.value = selectedPeriod.endTime.slice(0, 10);
    const progress = [
      shadow.getElementById('collector-live'),
      shadow.getElementById('collector-product'),
    ];
    const start = shadow.getElementById('collector-start');
    const stop = shadow.getElementById('collector-stop');
    const download = shadow.getElementById('collector-download');
    const status = shadow.getElementById('collector-status');
    const count = shadow.getElementById('collector-count');
    const rows = shadow.getElementById('collector-rows');
    const table = rows.closest('.collector-table-scroll');
    let lastRun;
    let lastMatchCount;

    function renderPreview(run, matches) {
      if (lastRun === run && lastMatchCount === matches) return;
      lastRun = run;
      lastMatchCount = matches;
      const preview = [];
      for (const scan of run?.scans ?? []) {
        for (const item of scan.matches.values()) {
          if (preview.length >= previewLimit) break;
          const tr = document.createElement('tr');
          const values = [
            modes[scan.mode].label, item.name, numberFormat.format(item.spendCents / 100),
            numberFormat.format(item.impressions),
          ];
          for (const value of values) {
            const td = document.createElement('td');
            td.textContent = value;
            tr.append(td);
          }
          preview.push(tr);
        }
      }
      rows.replaceChildren(...preview);
      table.hidden = preview.length === 0;
    }

    function renderPanel() {
      const list = templatesForPage();
      const ready = Boolean(list[0] && list[1]);
      const run = batch?.account === currentPage() ? batch : null;
      const running = batch?.phase === 'running';
      const complete = run?.phase === 'complete';
      if (periodAutomatic && !running) {
        dateStart.value = selectedPeriod.startTime.slice(0, 10);
        dateEnd.value = selectedPeriod.endTime.slice(0, 10);
      }
      dateStart.disabled = running;
      dateEnd.disabled = running;
      month.disabled = running;
      const scope = run ?? selectedPeriod;
      const differentScopes = list.filter((template) => template && !samePeriod(template, scope));
      scopeNote.hidden = differentScopes.length === 0;
      scopeNote.textContent = scopeNote.hidden ? '' : [
        ...differentScopes.map((template) => `${modes[template.mode].label} 页面查询 ${periodText(template)}`),
        `采集日期 ${periodText(scope)}`,
        '核对采集序号时，请将对应页面设为相同日期，并按展现从高到低排序。',
      ].join('。');

      for (const [index, mode] of modeCodes.entries()) {
        progress[index].textContent = scanProgress(mode, list[index], run?.scans[index]);
      }
      start.hidden = running;
      start.disabled = !ready || Boolean(dateError);
      start.textContent = run ? (canResume(run, list) ? '继续采集' : '重新采集') : '开始采集';
      stop.hidden = !running;
      stop.disabled = Boolean(batch?.pauseRequested);
      download.disabled = !complete || Boolean(dateError);

      if (dateError) {
        status.textContent = dateError;
      } else if (running) {
        status.textContent = batch.pauseRequested ? '正在暂停' : '正在逐页采集，请保持页面打开。';
      } else if (complete) {
        status.textContent = '两类素材筛选完成，可以下载汇总表。';
      } else if (run?.phase === 'paused') {
        status.textContent = run.error || '采集已暂停，进度已保留。';
      } else {
        status.textContent = ready ? '' : '请分别打开推直播间和推商品，等待捕获列表请求。';
      }

      const totals = summarize(run);
      count.textContent = [
        `已采集 ${numberFormat.format(totals.scanned)} 条`,
        `达标 ${numberFormat.format(totals.matches)} 条`,
        `消耗 ${numberFormat.format(totals.spendCents / 100)} 元`,
        `展现 ${numberFormat.format(totals.impressions)} 次`,
        totals.matches > previewLimit ? `预览前 ${previewLimit} 条` : '',
      ].filter(Boolean).join(' ');
      renderPreview(run, totals.matches);
    }

    function setPanelOpen(open) {
      panel.hidden = !open;
      openButton.hidden = open;
      openButton.setAttribute('aria-expanded', String(open));
      if (open) {
        renderPanel();
        (!start.hidden && !start.disabled ? start : closeButton).focus();
      } else {
        openButton.focus();
      }
    }

    openButton.addEventListener('click', () => setPanelOpen(true));
    closeButton.addEventListener('click', () => setPanelOpen(false));
    function applyDates(automatic = false) {
      if (batch?.phase === 'running') return;
      periodAutomatic = automatic === true;
      const startDate = dateStart.value;
      const endDate = dateEnd.value;
      dateError = '';
      if (!isDate(startDate) || !isDate(endDate) || startDate > endDate) {
        dateError = '请选择有效的起止日期，结束日期须等于或晚于开始日期。';
      } else {
        const nextPeriod = { startTime: `${startDate} 00:00:00`, endTime: `${endDate} 23:59:59` };
        if (selectedPeriod.startTime !== nextPeriod.startTime || selectedPeriod.endTime !== nextPeriod.endTime) {
          selectedPeriod = nextPeriod;
          batch = undefined;
        }
      }
      renderPanel();
    }
    dateStart.addEventListener('input', () => applyDates());
    dateEnd.addEventListener('input', () => applyDates());
    month.addEventListener('click', () => {
      const currentMonth = recentPageMonthPeriod(templatesForPage());
      dateStart.value = currentMonth.startTime.slice(0, 10);
      dateEnd.value = currentMonth.endTime.slice(0, 10);
      applyDates(true);
    });
    panel.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') setPanelOpen(false);
    });
    start.addEventListener('click', startCollection);
    stop.addEventListener('click', () => {
      if (batch?.phase !== 'running') return;
      batch.pauseRequested = true;
      renderPanel();
    });
    download.addEventListener('click', () => {
      if (batch?.account === currentPage() && batch.phase === 'complete') downloadCsv(batch);
    });

    document.body.append(host);
    updatePanel = () => {
      host.hidden = currentPage() === null;
      if (!host.hidden && !panel.hidden) renderPanel();
    };
    window.addEventListener('popstate', updatePanel);
    window.navigation?.addEventListener('currententrychange', updatePanel);
    updatePanel();
    return () => {
      if (currentPage() !== null) setPanelOpen(true);
    };
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
  const Xhr = window.XMLHttpRequest;
  const xhrPrototype = Xhr.prototype;
  const xhrOpen = xhrPrototype.open;
  const xhrSetHeader = xhrPrototype.setRequestHeader;
  const xhrSend = xhrPrototype.send;

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
      const xhr = new window.XMLHttpRequest();
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

  function syncAutomaticPeriod() {
    if (!periodAutomatic || batch?.phase === 'running') return;
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

  xhrPrototype.open = function open(method, input) {
    const result = xhrOpen.apply(this, arguments);
    requests.delete(this);
    if (replayRequests.has(this)) return result;
    try {
      const url = onSelfVideoPage() ? matchEndpoint(method, input) : null;
      if (url) requests.set(this, { url, headers: [] });
    } catch {
      // Leave invalid requests to the page's original transport.
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
    if (request && onSelfVideoPage()) {
      request.sequence = nextSequence++;
      const candidate = readTemplate(request, body, this.withCredentials);
      if (candidate) {
        listener = () => saveTemplate(candidate, readXhr(this));
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
    const run = canResume(batch) ? batch : planBatch();
    batch = run;
    run.phase = 'running';
    run.pauseRequested = false;
    run.error = '';
    updatePanel();
    try {
      await collect(run);
    } catch (error) {
      run.error = error.message;
    } finally {
      run.phase = !run.restartRequired && run.complete ? 'complete' : 'paused';
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

  function panelMarkup() {
    return `
      <style>
        :host {
          --self-ink: #1d293a;
          --self-muted: #516176;
          --self-border: #dbe2eb;
          --self-blue: #1755c7;
          --self-tint: #f4f7fb;
          font-family: inherit;
          font-size: 13px;
          line-height: 1.45;
        }
        *, *::before, *::after { box-sizing: border-box; }
        [hidden] { display: none !important; }
        button, input { font: inherit; }
        button { cursor: pointer; }
        .self-launcher, .self-panel {
          position: fixed;
          z-index: 2147483600;
          inset-inline-end: 18px;
          inset-block-end: 18px;
          color: var(--self-ink);
        }
        button {
          min-height: 34px;
          padding: 5px 12px;
          color: var(--self-ink);
          background: #fff;
          border: 1px solid var(--self-border);
          border-radius: 6px;
        }
        button:hover:not(:disabled) { background: var(--self-tint); }
        button:disabled { color: #687689; opacity: 0.7; cursor: default; }
        button:focus-visible, input:focus-visible { outline: 2px solid var(--self-blue); outline-offset: 2px; }
        .self-primary { color: #fff; background: var(--self-blue); border-color: var(--self-blue); font-weight: 600; }
        .self-primary:hover:not(:disabled) { color: #fff; background: #1248ad; }
        .self-launcher { min-height: 40px; padding-inline: 16px; box-shadow: 0 3px 12px rgb(28 49 83 / 20%); }
        .self-panel {
          width: min(480px, calc(100vw - 36px));
          max-height: min(620px, calc(100dvh - 36px));
          overflow: auto;
          overscroll-behavior: contain;
          background: #fff;
          border: 1px solid var(--self-border);
          border-radius: 10px;
          box-shadow: 0 10px 30px rgb(24 42 68 / 18%);
        }
        .self-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 12px 16px; border-bottom: 1px solid var(--self-border); }
        .self-header h2 { margin: 0; font-size: 16px; font-weight: 650; }
        .self-close { min-height: 30px; padding: 3px 9px; }
        .self-body { display: grid; gap: 12px; padding: 14px 16px 16px; }
        .self-period-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
        .self-period, .self-status, .self-count { margin: 0; }
        .self-period, .self-status { color: var(--self-muted); overflow-wrap: anywhere; }
        .self-status:empty { display: none; }
        .self-date-fields { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
        .self-date-fields label { display: grid; gap: 6px; min-width: 0; }
        .self-date-fields input { width: 100%; min-width: 0; height: 34px; padding: 4px 8px; color: var(--self-ink); background: #fff; border: 1px solid var(--self-border); border-radius: 6px; }
        .self-date-fields input:disabled { color: var(--self-muted); background: var(--self-tint); }
        .self-progress { min-height: 32px; margin: 0; padding: 6px 8px; background: var(--self-tint); border-radius: 5px; font-variant-numeric: tabular-nums; overflow-wrap: anywhere; }
        .self-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
        .self-note { padding: 8px 10px; color: #7a4f00; background: #fff8e6; border-inline-start: 3px solid #e5a400; border-radius: 5px; }
        .self-details { display: grid; gap: 8px; padding-top: 12px; border-top: 1px solid var(--self-border); }
        .self-filter { margin: 0; color: var(--self-muted); }
        .self-count { color: var(--self-ink); font-weight: 600; }
        .self-table-scroll { max-height: 230px; overflow: auto; overscroll-behavior: contain; border: 1px solid var(--self-border); border-radius: 6px; }
        table { width: 100%; table-layout: fixed; border-spacing: 0; text-align: left; }
        th, td { padding: 7px 8px; vertical-align: top; overflow-wrap: anywhere; border-bottom: 1px solid var(--self-border); }
        th { position: sticky; inset-block-start: 0; background: var(--self-tint); color: var(--self-muted); font-weight: 650; white-space: nowrap; }
        th:last-child, td:last-child { width: 30%; text-align: right; font-variant-numeric: tabular-nums; }
        tbody tr:last-child td { border-bottom: 0; }
        @media (max-width: 400px) {
          .self-launcher, .self-panel { inset-inline-end: 12px; inset-block-end: 12px; }
          .self-panel { width: calc(100vw - 24px); max-height: calc(100dvh - 24px); }
          .self-header, .self-body { padding-inline: 12px; }
        }
      </style>
      <button id="self-open" class="self-launcher self-primary" type="button" aria-controls="self-panel" aria-expanded="false">自营视频采集</button>
      <section id="self-panel" class="self-panel" aria-labelledby="self-title" hidden>
        <header class="self-header">
          <h2 id="self-title">自营视频采集</h2>
          <button id="self-close" class="self-close" type="button" aria-label="关闭自营视频采集面板">关闭</button>
        </header>
        <div class="self-body">
          <div class="self-period-row">
            <p class="self-period">采集日期</p>
            <button id="self-month" type="button" title="本月 1 日至最近页面查询截止日，无近期查询时截至昨天。每月 1 日选取上月。">本月</button>
          </div>
          <div class="self-date-fields">
            <label>开始日期<input id="self-date-start" type="date" required></label>
            <label>结束日期<input id="self-date-end" type="date" required></label>
          </div>
          <p id="self-scope-note" class="self-status self-note" role="status" hidden></p>
          <p id="self-progress" class="self-progress">自营视频 待捕获</p>
          <div class="self-actions">
            <button id="self-start" class="self-primary" type="button" disabled>开始采集</button>
            <button id="self-stop" type="button" hidden>暂停采集</button>
            <button id="self-download" type="button" disabled>下载 CSV</button>
          </div>
          <p id="self-status" class="self-status" role="status"></p>
          <div class="self-details">
            <p class="self-filter">观看次数 &gt; ${numberFormat.format(minWatchCount)} 次</p>
            <p id="self-count" class="self-count">已采集 0 条 达标 0 条</p>
            <div class="self-table-scroll" hidden>
              <table aria-label="符合筛选条件的自营视频">
                <thead><tr><th scope="col">标题</th><th scope="col">观看次数</th></tr></thead>
                <tbody id="self-rows"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
    `;
  }

  function mountPanel() {
    const host = document.createElement('div');
    host.id = 'compass-self-video-collector';
    const shadow = host.attachShadow({ mode: 'open' });
    shadow.innerHTML = panelMarkup();
    const openButton = shadow.getElementById('self-open');
    const panel = shadow.getElementById('self-panel');
    const closeButton = shadow.getElementById('self-close');
    const dateStart = shadow.getElementById('self-date-start');
    const dateEnd = shadow.getElementById('self-date-end');
    const month = shadow.getElementById('self-month');
    const scopeNote = shadow.getElementById('self-scope-note');
    const progress = shadow.getElementById('self-progress');
    const start = shadow.getElementById('self-start');
    const stop = shadow.getElementById('self-stop');
    const download = shadow.getElementById('self-download');
    const status = shadow.getElementById('self-status');
    const count = shadow.getElementById('self-count');
    const rows = shadow.getElementById('self-rows');
    const table = rows.closest('.self-table-scroll');
    let lastRun;
    let lastMatchCount;

    dateStart.value = selectedPeriod.startDate;
    dateEnd.value = selectedPeriod.endDate;

    function renderPreview(run) {
      const matchCount = run?.matches.size ?? 0;
      if (lastRun === run && lastMatchCount === matchCount) return;
      lastRun = run;
      lastMatchCount = matchCount;
      const preview = sortedMatches(run).slice(0, previewLimit).map((item) => {
        const tr = document.createElement('tr');
        for (const value of [item.title, numberFormat.format(item.watchCount)]) {
          const td = document.createElement('td');
          td.textContent = value;
          tr.append(td);
        }
        return tr;
      });
      rows.replaceChildren(...preview);
      table.hidden = preview.length === 0;
    }

    function renderPanel() {
      const run = batch;
      const running = run?.phase === 'running';
      const complete = run?.phase === 'complete';
      if (periodAutomatic && !running) {
        dateStart.value = selectedPeriod.startDate;
        dateEnd.value = selectedPeriod.endDate;
      }
      dateStart.disabled = running;
      dateEnd.disabled = running;
      month.disabled = running;

      const scope = run ?? selectedPeriod;
      const differentScope = template && !samePeriod(template.sourcePeriod, scope);
      scopeNote.hidden = !differentScope;
      scopeNote.textContent = differentScope
        ? `最近页面查询 ${periodText(template.sourcePeriod)}。采集日期 ${periodText(scope)}。核对序号时，请将页面设为相同日期并按观看次数从高到低排序。`
        : '';
      progress.textContent = progressText(run);
      start.hidden = running;
      start.disabled = !template || Boolean(dateError);
      start.textContent = run ? (canResume(run) ? '继续采集' : '重新采集') : '开始采集';
      stop.hidden = !running;
      stop.disabled = Boolean(run?.pauseRequested);
      download.disabled = !complete || Boolean(dateError);

      if (dateError) status.textContent = dateError;
      else if (running) status.textContent = run.pauseRequested ? '正在暂停' : '正在逐页采集，请保持页面打开。';
      else if (complete) status.textContent = '自营视频筛选完成，可以下载结果。';
      else if (run?.phase === 'paused') status.textContent = run.error || '采集已暂停，进度已保留。';
      else status.textContent = template ? '' : '请在视频明细列表查询一次，等待捕获请求。';

      const matchCount = run?.matches.size ?? 0;
      count.textContent = [
        `已采集 ${numberFormat.format(run?.scanned ?? 0)} 条`,
        `达标 ${numberFormat.format(matchCount)} 条`,
        matchCount > previewLimit ? `预览前 ${previewLimit} 条` : '',
      ].filter(Boolean).join(' ');
      renderPreview(run);
    }

    function setPanelOpen(open) {
      panel.hidden = !open;
      openButton.hidden = open;
      openButton.setAttribute('aria-expanded', String(open));
      if (open) {
        renderPanel();
        (!start.hidden && !start.disabled ? start : closeButton).focus();
      } else {
        openButton.focus();
      }
    }

    function applyDates(automatic = false) {
      if (batch?.phase === 'running') return;
      periodAutomatic = automatic === true;
      const startDate = dateStart.value;
      const endDate = dateEnd.value;
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
      renderPanel();
    }

    openButton.addEventListener('click', () => setPanelOpen(true));
    closeButton.addEventListener('click', () => setPanelOpen(false));
    dateStart.addEventListener('input', () => applyDates());
    dateEnd.addEventListener('input', () => applyDates());
    month.addEventListener('click', () => {
      const current = recentPageMonthPeriod(template?.sourcePeriod);
      dateStart.value = current.startDate;
      dateEnd.value = current.endDate;
      applyDates(true);
    });
    panel.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') setPanelOpen(false);
    });
    start.addEventListener('click', startCollection);
    stop.addEventListener('click', () => {
      if (batch?.phase !== 'running') return;
      batch.pauseRequested = true;
      renderPanel();
    });
    download.addEventListener('click', () => {
      if (batch?.phase === 'complete') downloadCsv(batch);
    });

    document.body.append(host);
    updatePanel = () => {
      host.hidden = !onSelfVideoPage();
      if (!host.hidden && !panel.hidden) renderPanel();
    };
    window.addEventListener('popstate', updatePanel);
    window.navigation?.addEventListener('currententrychange', updatePanel);
    updatePanel();
    return () => {
      if (onSelfVideoPage()) setPanelOpen(true);
    };
  }

  let openPanel;
  if (document.body) openPanel = mountPanel();
  else document.addEventListener('DOMContentLoaded', () => { openPanel = mountPanel(); }, { once: true });
  GM_registerMenuCommand('打开自营视频采集面板', () => openPanel?.());
})();
