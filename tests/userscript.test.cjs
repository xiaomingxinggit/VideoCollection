const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const userscriptPath = path.join(__dirname, '..', 'qianchuan-material-collector.user.js');
const userscriptSource = fs.readFileSync(userscriptPath, 'utf8');

class FakeXhr {
  open() {}
  setRequestHeader() {}
  send() {}
  abort() {}
  addEventListener() {}
  removeEventListener() {}
}

function loadQianchuanModule() {
  const originMarker = "const origin = 'https://qianchuan.jinritemai.com';";
  const originIndex = userscriptSource.indexOf(originMarker);
  const opening = userscriptSource.lastIndexOf('(() => {', originIndex);
  const closing = userscriptSource.indexOf('\n})();', opening);
  assert.notEqual(opening, -1, 'Qianchuan IIFE opening marker must exist');
  assert.notEqual(closing, -1, 'Qianchuan IIFE closing marker must exist');
  const exports = `
    globalThis.__qianchuanTest = { csvText, summarize, workbookSnapshot };
  `;
  const source = `${userscriptSource.slice(opening, closing)}\n${exports}${userscriptSource.slice(closing, closing + '\n})();'.length)}`;
  const sandbox = {
    AbortController,
    Blob,
    Date,
    Intl,
    Map,
    Object,
    Request,
    Set,
    URL,
    URLSearchParams,
    WeakMap,
    WeakSet,
    clearTimeout,
    console,
    document: { body: null, addEventListener() {} },
    fetch() {},
    GM_registerMenuCommand() {},
    location: {
      href: 'https://qianchuan.jinritemai.com/dataV3/roi2-material-analysis?aavid=advertiser-a',
      origin: 'https://qianchuan.jinritemai.com',
      pathname: '/dataV3/roi2-material-analysis',
      search: '?aavid=advertiser-a',
    },
    setTimeout,
  };
  sandbox.window = sandbox;
  sandbox.XMLHttpRequest = FakeXhr;
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox, { filename: userscriptPath });
  return sandbox.__qianchuanTest;
}

function loadCompassModule() {
  const source = userscriptSource;
  const closing = source.lastIndexOf('\n})();');
  assert.notEqual(closing, -1, 'Compass IIFE closing marker must exist');
  const exports = `
    globalThis.__compassTest = {
      addPage,
      csvText,
      currentMonthPeriod,
      readPage,
      recentPageMonthPeriod,
      replayUrl,
      rowData,
    };
  `;
  const instrumented = `${source.slice(0, closing)}\n${exports}${source.slice(closing)}`;
  const sandbox = {
    AbortController,
    Blob,
    Date,
    Intl,
    Map,
    Object,
    Set,
    URL,
    WeakMap,
    WeakSet,
    clearTimeout,
    console,
    document: { body: null, addEventListener() {} },
    GM_registerMenuCommand() {},
    location: {
      href: 'https://compass.jinritemai.com/shop/video/self',
      origin: 'https://compass.jinritemai.com',
      pathname: '/shop/video/self',
      search: '',
    },
    setTimeout,
  };
  sandbox.window = sandbox;
  sandbox.XMLHttpRequest = FakeXhr;
  vm.createContext(sandbox);
  vm.runInContext(instrumented, sandbox, { filename: userscriptPath });
  return sandbox.__compassTest;
}

function loadSharedUiModule() {
  const opening = userscriptSource.indexOf('const videoCollectionLiquidUi = (() => {');
  const closing = userscriptSource.indexOf('\n})();', opening);
  assert.notEqual(opening, -1, 'Shared UI IIFE opening marker must exist');
  assert.notEqual(closing, -1, 'Shared UI IIFE closing marker must exist');
  const source = userscriptSource.slice(opening, closing + '\n})();'.length);
  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(`${source}\nglobalThis.__sharedUiTest = videoCollectionLiquidUi;`, sandbox, {
    filename: userscriptPath,
  });
  return sandbox.__sharedUiTest;
}

function loadWorkbookModule() {
  const opening = userscriptSource.indexOf('const videoCollectionWorkbook = (() => {');
  const closing = userscriptSource.indexOf('\n})();', opening);
  assert.notEqual(opening, -1, 'Workbook IIFE opening marker must exist');
  assert.notEqual(closing, -1, 'Workbook IIFE closing marker must exist');
  const source = userscriptSource.slice(opening, closing + '\n})();'.length);
  const store = new Map();
  const downloads = [];
  const listeners = new Map();
  let nextListener = 1;
  const encodeCell = ({ r, c }) => `${String.fromCharCode(65 + c)}${r + 1}`;
  const XLSX = {
    utils: {
      encode_cell: encodeCell,
      book_new: () => ({ SheetNames: [], Sheets: {} }),
      aoa_to_sheet(rows) {
        const sheet = { __rows: rows };
        const columns = Math.max(1, ...rows.map((row) => row.length));
        sheet['!ref'] = `A1:${encodeCell({ r: Math.max(0, rows.length - 1), c: columns - 1 })}`;
        rows.forEach((row, r) => row.forEach((value, c) => {
          sheet[encodeCell({ r, c })] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
        }));
        return sheet;
      },
      book_append_sheet(workbook, sheet, name) {
        workbook.SheetNames.push(name);
        workbook.Sheets[name] = sheet;
      },
    },
    writeFileXLSX(workbook, filename, options) {
      downloads.push({ workbook, filename, options });
    },
  };
  const sandbox = {
    Date,
    Intl,
    XLSX,
    GM_getValue: (key, fallback) => store.has(key) ? store.get(key) : fallback,
    GM_setValue: (key, value) => store.set(key, value),
    GM_deleteValue: (key) => store.delete(key),
    GM_addValueChangeListener(key, callback) {
      const id = nextListener++;
      listeners.set(id, { callback, key });
      return id;
    },
    GM_removeValueChangeListener: (id) => listeners.delete(id),
  };
  vm.createContext(sandbox);
  vm.runInContext(`${source}\nglobalThis.__workbookTest = videoCollectionWorkbook;`, sandbox, {
    filename: userscriptPath,
  });
  return { downloads, listeners, module: sandbox.__workbookTest, store };
}

function videoRow(id, title, watchCount) {
  return {
    cell_info: {
      video: { video: { video_id: id, video_title: title } },
      watch_cnt: { index_values: { value: { value: watchCount } } },
    },
  };
}

function runState(pageSize = 3) {
  return {
    pageNo: 1,
    pageSize,
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

const compass = loadCompassModule();
const qianchuan = loadQianchuanModule();
const sharedUi = loadSharedUiModule();
const workbook = loadWorkbookModule();

test('userscript pins Vue, Element Plus, and SheetJS with SUN and Dark themes', () => {
  const vueBridge = 'data:text/javascript;base64,Y29uc3QgX192Y0dsb2JhbCA9IGdsb2JhbFRoaXM7IGNvbnN0IF9fdmNQYWdlR2xvYmFsID0gdHlwZW9mIHVuc2FmZVdpbmRvdyA9PT0gJ29iamVjdCcgPyB1bnNhZmVXaW5kb3cgOiBfX3ZjR2xvYmFsOyBjb25zdCBfX3ZjSGFkVnVlID0gT2JqZWN0Lmhhc093bihfX3ZjR2xvYmFsLCAnVnVlJyk7IGNvbnN0IF9fdmNPbGRWdWUgPSBfX3ZjR2xvYmFsLlZ1ZTsgY29uc3QgX192Y0hhZEVsZW1lbnRQbHVzID0gT2JqZWN0Lmhhc093bihfX3ZjR2xvYmFsLCAnRWxlbWVudFBsdXMnKTsgY29uc3QgX192Y09sZEVsZW1lbnRQbHVzID0gX192Y0dsb2JhbC5FbGVtZW50UGx1czsgY29uc3QgX192Y0hhZFBhZ2VWdWUgPSBPYmplY3QuaGFzT3duKF9fdmNQYWdlR2xvYmFsLCAnVnVlJyk7IGNvbnN0IF9fdmNPbGRQYWdlVnVlID0gX192Y1BhZ2VHbG9iYWwuVnVlOyBjb25zdCBfX3ZjSGFkUGFnZUVsZW1lbnRQbHVzID0gT2JqZWN0Lmhhc093bihfX3ZjUGFnZUdsb2JhbCwgJ0VsZW1lbnRQbHVzJyk7IGNvbnN0IF9fdmNPbGRQYWdlRWxlbWVudFBsdXMgPSBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1czsgX192Y0dsb2JhbC5WdWUgPSBWdWU7IF9fdmNQYWdlR2xvYmFsLlZ1ZSA9IFZ1ZTs=';
  const elementBridge = 'data:text/javascript;base64,Y29uc3QgRWxlbWVudFBsdXMgPSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzOyBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1cyA9IEVsZW1lbnRQbHVzOyBjb25zdCBfX3ZjUmVzdG9yZUZyYW1ld29ya0dsb2JhbHMgPSAoKSA9PiB7IGlmIChfX3ZjSGFkVnVlKSBfX3ZjR2xvYmFsLlZ1ZSA9IF9fdmNPbGRWdWU7IGVsc2UgZGVsZXRlIF9fdmNHbG9iYWwuVnVlOyBpZiAoX192Y0hhZEVsZW1lbnRQbHVzKSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzID0gX192Y09sZEVsZW1lbnRQbHVzOyBlbHNlIGRlbGV0ZSBfX3ZjR2xvYmFsLkVsZW1lbnRQbHVzOyBpZiAoX192Y0hhZFBhZ2VWdWUpIF9fdmNQYWdlR2xvYmFsLlZ1ZSA9IF9fdmNPbGRQYWdlVnVlOyBlbHNlIGRlbGV0ZSBfX3ZjUGFnZUdsb2JhbC5WdWU7IGlmIChfX3ZjSGFkUGFnZUVsZW1lbnRQbHVzKSBfX3ZjUGFnZUdsb2JhbC5FbGVtZW50UGx1cyA9IF9fdmNPbGRQYWdlRWxlbWVudFBsdXM7IGVsc2UgZGVsZXRlIF9fdmNQYWdlR2xvYmFsLkVsZW1lbnRQbHVzOyB9Ow==';
  const dependencies = [
    'https://cdn.jsdelivr.net/npm/vue@3.5.46/dist/vue.global.prod.js#sha256-wuS4iOy7EseyT9Xov2VIPOqrc8N7eUAveWgCOXftPoE=',
    'https://cdn.jsdelivr.net/npm/element-plus@2.14.8/dist/index.full.min.js#sha256-D/7wpbpse4CCbLyQwzY2KGIjP5bvtkEHnzgXXJyOebI=',
    'https://cdn.jsdelivr.net/npm/element-plus@2.14.8/dist/index.css#sha256-9vVHqUarXedNRDZ6BVb0lN1jDGdZmtT2amrCAKhegKY=',
    'https://cdn.jsdelivr.net/npm/element-plus@2.14.8/theme-chalk/dark/css-vars.css#sha256-KOclPpHatx0lmXDXwrIWdM7/op+kRvWWikCTMhUWFgA=',
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.mini.min.js#sha256-DLNT+DDXKIOFSSyD0newWN3qxmTKUc8Tk6of0+K3CTk=',
  ];
  for (const dependency of dependencies) assert.ok(userscriptSource.includes(dependency));
  assert.match(userscriptSource, /@version\s+1\.6\.0/);
  const vueIndex = userscriptSource.indexOf(dependencies[0]);
  const vueBridgeIndex = userscriptSource.indexOf(vueBridge);
  const elementIndex = userscriptSource.indexOf(dependencies[1]);
  const elementBridgeIndex = userscriptSource.indexOf(elementBridge);
  assert.ok(vueIndex < vueBridgeIndex && vueBridgeIndex < elementIndex && elementIndex < elementBridgeIndex);
  assert.match(userscriptSource, /Vue\.createApp/);
  assert.match(userscriptSource, /<el-segmented/);
  assert.match(userscriptSource, /<table class="vc-result-table"/);
  assert.match(userscriptSource, /scope="col"/);
  assert.match(userscriptSource, /startButton && !startButton\.disabled/);
  assert.match(userscriptSource, /GM_getValue\(themeKey, 'sun'\)/);
  assert.match(userscriptSource, /\{ label: 'SUN', value: 'sun' \}/);
  assert.match(userscriptSource, /\{ label: 'Dark', value: 'dark' \}/);
  assert.doesNotMatch(userscriptSource, /function panelMarkup/);
});

test('shared panel removes the launcher logo and supports drag with edge snapping', () => {
  assert.doesNotMatch(userscriptSource, /vc-launcher-orb/);
  assert.match(userscriptSource, /@pointerdown="startDrag\('launcher', \$event\)"/);
  assert.match(userscriptSource, /<header class="vc-header" @pointerdown="startDrag\('panel', \$event\)">/);
  assert.match(userscriptSource, /video-collection-liquid-position/);
  assert.match(userscriptSource, />下载总 Excel<\/el-button>/);
  assert.match(userscriptSource, /v-for="item in view\.workbookItems"/);
  assert.match(userscriptSource, /options\.subscribe\?\.\(\(\) => this\.sync\(\)\)/);

  const fit = (position) => JSON.parse(JSON.stringify(sharedUi.fitFloatingPosition(
    position,
    { width: 200, height: 100 },
    { width: 1000, height: 800 },
  )));
  assert.deepEqual(fit({ x: 20, y: 30 }), { x: 12, y: 12 });
  assert.deepEqual(fit({ x: 780, y: 680 }), { x: 788, y: 688 });
  assert.deepEqual(fit({ x: 400, y: 350 }), { x: 400, y: 350 });
  assert.deepEqual(fit({ x: 2000, y: -100 }), { x: 788, y: 12 });
});

test('userscript bridges Vue into Element Plus without leaking page globals', () => {
  const bridges = [...userscriptSource.matchAll(/@require\s+(data:text\/javascript;base64,\S+)/g)]
    .map((match) => Buffer.from(match[1].split(',')[1], 'base64').toString('utf8'))
    .filter((source) => source.includes('__vcGlobal'));
  assert.equal(bridges.length, 2);

  const pageVue = { page: 'vue' };
  const pageElementPlus = { page: 'element-plus' };
  const sandboxVue = { sandbox: 'vue' };
  const sandboxElementPlus = { sandbox: 'element-plus' };
  const pageWindow = { Vue: pageVue, ElementPlus: pageElementPlus };
  const sandbox = {
    pageVue,
    pageElementPlus,
    Vue: sandboxVue,
    ElementPlus: sandboxElementPlus,
    unsafeWindow: pageWindow,
  };
  vm.createContext(sandbox);
  vm.runInContext(`
    (() => {
      var Vue = { runtime: 'vue' };
      ${bridges[0]}
      globalThis.ElementPlus = { runtime: 'element-plus', vue: globalThis.Vue };
      ${bridges[1]}
      globalThis.__bridgeBeforeCleanup = {
        Vue: globalThis.Vue,
        ElementPlus: globalThis.ElementPlus,
        pageVue: unsafeWindow.Vue,
        pageElementPlus: unsafeWindow.ElementPlus,
      };
      globalThis.__bridgeResult = { Vue, ElementPlus };
      __vcRestoreFrameworkGlobals();
    })();
  `, sandbox);

  assert.equal(sandbox.__bridgeResult.ElementPlus.vue, sandbox.__bridgeResult.Vue);
  assert.equal(sandbox.__bridgeBeforeCleanup.Vue, sandbox.__bridgeResult.Vue);
  assert.equal(sandbox.__bridgeBeforeCleanup.ElementPlus, sandbox.__bridgeResult.ElementPlus);
  assert.equal(sandbox.__bridgeBeforeCleanup.pageVue, sandbox.__bridgeResult.Vue);
  assert.equal(sandbox.__bridgeBeforeCleanup.pageElementPlus, sandbox.__bridgeResult.ElementPlus);
  assert.equal(sandbox.Vue, sandboxVue);
  assert.equal(sandbox.ElementPlus, sandboxElementPlus);
  assert.equal(pageWindow.Vue, pageVue);
  assert.equal(pageWindow.ElementPlus, pageElementPlus);
});

test('userscript loads SheetJS between bridges and restores the page global', () => {
  const sheetDependency = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.mini.min.js#sha256-DLNT+DDXKIOFSSyD0newWN3qxmTKUc8Tk6of0+K3CTk=';
  const sheetBridges = [...userscriptSource.matchAll(/@require\s+(data:text\/javascript;base64,\S+)/g)]
    .map((match) => ({
      encoded: match[1],
      source: Buffer.from(match[1].split(',')[1], 'base64').toString('utf8'),
    }))
    .filter((item) => item.source.includes('__vcSheetPageGlobal'));
  assert.equal(sheetBridges.length, 2);
  const beforeIndex = userscriptSource.indexOf(sheetBridges[0].encoded);
  const libraryIndex = userscriptSource.indexOf(sheetDependency);
  const afterIndex = userscriptSource.indexOf(sheetBridges[1].encoded);
  assert.ok(beforeIndex < libraryIndex && libraryIndex < afterIndex);

  const pageXlsx = { page: 'xlsx' };
  const sandbox = { pageXlsx, unsafeWindow: { XLSX: pageXlsx } };
  vm.createContext(sandbox);
  vm.runInContext(`
    (() => {
      ${sheetBridges[0].source}
      var XLSX = { runtime: 'xlsx' };
      unsafeWindow.XLSX = XLSX;
      ${sheetBridges[1].source}
      globalThis.__sheetBridgeResult = { local: XLSX, page: unsafeWindow.XLSX };
    })();
  `, sandbox);
  assert.equal(sandbox.__sheetBridgeResult.local.runtime, 'xlsx');
  assert.equal(sandbox.__sheetBridgeResult.page, pageXlsx);
});

test('shared storage combines matching page snapshots into a two-sheet workbook', () => {
  const qianchuan = {
    scope: 'advertiser-a',
    startDate: '2026-11-01',
    endDate: '2026-11-25',
    scanned: 200,
    rows: [{
      promotionType: '推直播间',
      materialId: 'm1',
      title: '=保留为标题',
      spendYuan: 1234.56,
      impressions: 45678,
      videoUrl: 'https://example.com/video.mp4?token=secret',
      headers: { cookie: 'secret' },
    }],
  };
  const compassSnapshot = {
    scope: 'shop-a',
    startDate: '2026-11-01',
    endDate: '2026-11-25',
    scanned: 100,
    rows: [{
      title: '自营视频', watchCount: 30001, signedUrl: 'https://example.com/?a_bogus=secret',
    }],
  };

  assert.equal(workbook.module.save('qianchuan', qianchuan), true);
  assert.equal(workbook.store.size, 1);
  assert.equal(workbook.module.summary().ready, false);
  assert.doesNotMatch(
    [...workbook.store.values()].join('\n'),
    /videoUrl|msToken|a_bogus|headers|advertiser-a|shop-a/,
  );
  assert.equal(workbook.module.save('compass', { ...compassSnapshot, endDate: '2026-11-24' }), true);
  assert.equal(workbook.store.size, 2);
  assert.equal(workbook.module.summary().ready, false);
  assert.match(workbook.module.summary().message, /日期不一致/);

  assert.equal(workbook.module.save('compass', compassSnapshot), true);
  const state = workbook.module.summary();
  assert.equal(state.ready, true);
  assert.deepEqual(JSON.parse(JSON.stringify(state.items.map((item) => item.state))), ['done', 'done']);
  const result = workbook.module.buildWorkbook(state);
  assert.deepEqual(result.SheetNames, ['千川素材', '自营视频']);
  assert.equal(result.Sheets['千川素材'].D2.v, 1234.56);
  assert.equal(result.Sheets['千川素材'].D2.t, 'n');
  assert.equal(result.Sheets['千川素材'].E2.v, 45678);
  assert.equal(result.Sheets['自营视频'].B2.v, 30001);
  assert.equal(result.Sheets['自营视频'].B2.t, 'n');

  const filename = workbook.module.downloadCombined();
  assert.equal(filename, 'video-collection-summary-2026-11-01_2026-11-25.xlsx');
  assert.equal(workbook.downloads.length, 1);
  assert.deepEqual(workbook.downloads[0].workbook.SheetNames, ['千川素材', '自营视频']);
  assert.deepEqual(JSON.parse(JSON.stringify(workbook.downloads[0].options)), { compression: true });
});

test('combined workbook rejects snapshots collected more than 24 hours apart', () => {
  const isolated = loadWorkbookModule();
  const period = { startDate: '2026-11-01', endDate: '2026-11-25' };
  assert.equal(isolated.module.save('qianchuan', {
    scope: 'advertiser-a', ...period, scanned: 0, rows: [],
  }), true);
  assert.equal(isolated.module.save('compass', {
    scope: 'shop-a', ...period, scanned: 0, rows: [],
  }), true);

  const key = 'video-collection:qianchuan:v1';
  const snapshot = JSON.parse(isolated.store.get(key));
  snapshot.collectedAt = new Date(Date.parse(snapshot.collectedAt) - (25 * 60 * 60 * 1000)).toISOString();
  isolated.store.set(key, JSON.stringify(snapshot));

  const state = isolated.module.summary();
  assert.equal(state.ready, false);
  assert.match(state.message, /超过 24 小时/);
  assert.throws(() => isolated.module.buildWorkbook(state), /超过 24 小时/);
});

test('Qianchuan snapshot maps both promotion types without persisting signed links', () => {
  const run = {
    account: 'advertiser-a',
    startTime: '2026-11-01 00:00:00',
    endTime: '2026-11-25 23:59:59',
    scans: [
      {
        mode: '2', offset: 10, spendCents: 123456, impressions: 45678,
        matches: new Map([['live', {
          id: 'live-1', name: '直播素材', spendCents: 123456, impressions: 45678,
          link: 'https://example.com/video.mp4?token=secret',
        }]]),
      },
      {
        mode: '1', offset: 8, spendCents: 200000, impressions: 50000,
        matches: new Map([['product', {
          id: 'product-1', name: '商品素材', spendCents: 200000, impressions: 50000,
          link: 'https://example.com/product.mp4?token=secret',
        }]]),
      },
    ],
  };
  const snapshot = JSON.parse(JSON.stringify(qianchuan.workbookSnapshot(run)));
  assert.equal(snapshot.scope, 'advertiser-a');
  assert.equal(snapshot.scanned, 18);
  assert.deepEqual(snapshot.rows.map((row) => row.promotionType), ['推直播间', '推商品']);
  assert.equal(snapshot.rows[0].spendYuan, 1234.56);
  assert.equal(snapshot.rows[0].impressions, 45678);
  assert.ok(snapshot.rows.every((row) => !Object.hasOwn(row, 'videoUrl')));
  assert.doesNotMatch(JSON.stringify(snapshot), /token=secret/);

  const csv = qianchuan.csvText(run);
  assert.match(csv, /"推直播间"/);
  assert.match(csv, /"推商品"/);
  assert.match(csv, /"合计"/);
});

test('both collectors persist completed snapshots before automatic period sync', () => {
  for (const source of ['qianchuan', 'compass']) {
    const saveCall = `videoCollectionWorkbook.save('${source}'`;
    const saveIndex = userscriptSource.indexOf(saveCall);
    assert.notEqual(saveIndex, -1);
    const phaseIndex = userscriptSource.lastIndexOf("run.phase =", saveIndex);
    const syncIndex = userscriptSource.indexOf('syncAutomaticPeriod();', saveIndex);
    assert.ok(phaseIndex < saveIndex && saveIndex < syncIndex);
  }
});

test('an empty but complete snapshot still counts as collected', () => {
  const emptyWorkbook = loadWorkbookModule();
  assert.equal(emptyWorkbook.module.save('qianchuan', {
    scope: 'advertiser-a', startDate: '2026-11-01', endDate: '2026-11-25', scanned: 0, rows: [],
  }), true);
  assert.equal(emptyWorkbook.module.save('compass', {
    scope: 'shop-a', startDate: '2026-11-01', endDate: '2026-11-25', scanned: 0, rows: [],
  }), true);
  const state = emptyWorkbook.module.summary();
  assert.equal(state.ready, true);
  assert.deepEqual(JSON.parse(JSON.stringify(state.items.map((item) => item.text))), [
    '已采集 0 条 2026-11-01 至 2026-11-25',
    '已采集 0 条 2026-11-01 至 2026-11-25',
  ]);
});

test('invalid snapshots are rejected and a new account starts a clean pair', () => {
  const isolated = loadWorkbookModule();
  const period = { startDate: '2026-11-01', endDate: '2026-11-25' };
  const qianchuan = {
    scope: 'advertiser-a', ...period, scanned: 1,
    rows: [{ promotionType: '推商品', materialId: 'm1', title: '素材', spendYuan: 1000, impressions: 20000 }],
  };
  const compassSnapshot = {
    scope: 'shop-a', ...period, scanned: 1,
    rows: [{ title: '视频', watchCount: 20001 }],
  };
  assert.equal(isolated.module.save('qianchuan', qianchuan), true);
  assert.equal(isolated.module.save('compass', compassSnapshot), true);
  assert.equal(isolated.module.summary().ready, true);

  assert.equal(isolated.module.begin('qianchuan', 'advertiser-b', period), true);
  assert.equal(isolated.module.summary().qianchuan, null);
  assert.equal(isolated.module.summary().compass, null);
  assert.equal(isolated.module.save('qianchuan', {
    ...qianchuan, scope: 'advertiser-b', startDate: '2026-02-30',
  }), false);
  assert.equal(isolated.module.read('qianchuan'), null);
  assert.equal(isolated.module.save('compass', {
    ...compassSnapshot, rows: [{ title: '坏数据', watchCount: '20001' }],
  }), false);
  assert.equal(isolated.module.read('compass'), null);
});

test('shared snapshot listeners refresh both page states and are removable', () => {
  const isolated = loadWorkbookModule();
  let calls = 0;
  const ids = isolated.module.subscribe(() => { calls++; });
  assert.equal(ids.length, 2);
  for (const listener of isolated.listeners.values()) listener.callback();
  assert.equal(calls, 2);
  isolated.module.unsubscribe(ids);
  assert.equal(isolated.listeners.size, 0);
});

test('Compass replay changes only the required query parameters', () => {
  const source = new URL('https://compass.jinritemai.com/compass_api/shop/video/overview/video_list'
    + '?begin_date=2026%2F11%2F18+00%3A00%3A00&end_date=2026%2F11%2F24+00%3A00%3A00'
    + '&date_type=21&page_no=7&page_size=10&sort_field=pay_amt&is_asc=true'
    + '&account_type=1&range_type=0&activity_id=&_lid=fake-lid'
    + '&msToken=fake-ms-token&a_bogus=fake-a-bogus&fp=fake-fp&verifyFp=fake-verify');
  const sourcePeriod = { startDate: '2026-11-18', endDate: '2026-11-24' };
  const next = compass.replayUrl(source, 3,
    { startDate: '2026-11-01', endDate: '2026-11-25' }, sourcePeriod);
  const changed = new Set(['begin_date', 'end_date', 'date_type', 'page_no', 'sort_field', 'is_asc']);
  const untouched = (url) => [...url.searchParams].filter(([key]) => !changed.has(key));

  assert.deepEqual(untouched(next), untouched(source));
  assert.equal(next.searchParams.get('begin_date'), '2026/11/01 00:00:00');
  assert.equal(next.searchParams.get('end_date'), '2026/11/25 00:00:00');
  assert.equal(next.searchParams.get('date_type'), '4');
  assert.equal(next.searchParams.get('page_no'), '3');
  assert.equal(next.searchParams.get('page_size'), '10');
  assert.equal(next.searchParams.get('sort_field'), 'watch_cnt');
  assert.equal(next.searchParams.get('is_asc'), 'false');
  assert.equal(next.searchParams.get('msToken'), 'fake-ms-token');
  assert.equal(next.searchParams.get('a_bogus'), 'fake-a-bogus');

  const sameDates = compass.replayUrl(source, 1, sourcePeriod, sourcePeriod);
  assert.equal(sameDates.searchParams.get('date_type'), '21');
});

test('Compass response parser reads the HAR table contract', () => {
  const rows = [videoRow('v1', '测试视频', 45678)];
  const page = compass.readPage({
    status: 200,
    json: {
      st: 0,
      data: {
        module_data: {
          core_data_0: {
            compass_general_table_value: {
              data: rows,
              page_result: { page_no: 1, page_size: 10, total: 2697 },
            },
          },
        },
      },
    },
  });

  assert.equal(page.total, 2697);
  const item = compass.rowData(rows[0]);
  assert.equal(item.id, 'v1');
  assert.equal(item.title, '测试视频');
  assert.equal(item.watchCount, 45678);
});

test('20,000 is excluded and stops a valid descending scan', () => {
  const run = runState();
  compass.addPage(run, {
    pageNo: 1,
    pageSize: 3,
    total: 5,
    rows: [videoRow('v1', '达标', 20001), videoRow('v2', '边界', 20000), videoRow('v3', '低于', 19999)],
  });

  assert.equal(run.matches.size, 1);
  assert.equal(run.firstOutside.position, 2);
  assert.equal(run.firstOutside.watchCount, 20000);
  assert.equal(run.complete, true);
});

test('out-of-order values disable the early boundary and force a full scan', () => {
  const run = runState();
  compass.addPage(run, {
    pageNo: 1,
    pageSize: 3,
    total: 5,
    rows: [videoRow('v1', '一', 30000), videoRow('v2', '二', 19000), videoRow('v3', '三', 25000)],
  });
  assert.equal(run.descending, false);
  assert.equal(run.firstOutside, null);
  assert.equal(run.complete, false);

  compass.addPage(run, {
    pageNo: 2,
    pageSize: 3,
    total: 5,
    rows: [videoRow('v4', '四', 18000), videoRow('v5', '五', 10000)],
  });
  assert.equal(run.complete, true);
  assert.equal(run.matches.size, 2);
});

test('month shortcut ends at yesterday and follows a recent page cutoff', () => {
  const current = compass.currentMonthPeriod(new Date('2026-11-26T04:00:00Z'));
  assert.equal(current.startDate, '2026-11-01');
  assert.equal(current.endDate, '2026-11-25');
  const recent = compass.recentPageMonthPeriod({ startDate: '2026-11-01', endDate: '2026-11-24' });
  assert.equal(recent.startDate, '2026-11-01');
  assert.equal(recent.endDate, '2026-11-24');
});

test('Compass CSV has only title and watch-count columns', () => {
  const run = runState();
  run.matches.set('v1', { title: '=公式标题', watchCount: 30001 });
  run.matches.set('v2', { title: '普通标题', watchCount: 40002 });
  const lines = compass.csvText(run).slice(1).split('\r\n');

  assert.equal(lines[0], '"标题","观看次数"');
  assert.equal(lines.length, 3);
  assert.equal(lines[1], '"普通标题","40002"');
  assert.equal(lines[2], '"\'=公式标题","30001"');
  assert.ok(lines.every((line) => (line.match(/","/g) ?? []).length === 1));
});
