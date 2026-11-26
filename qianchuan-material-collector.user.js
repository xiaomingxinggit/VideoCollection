// ==UserScript==
// @name         千川素材请求采集
// @namespace    VideoCollection
// @version      1.3.9
// @description  按展现排序采集推直播间和推商品的视频素材，筛选后汇总为 CSV。
//
// @match        https://qianchuan.jinritemai.com/dataV3/roi2-material-analysis*
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
