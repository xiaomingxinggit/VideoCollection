const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const userscriptPath = path.join(__dirname, '..', 'qianchuan-material-collector.user.js');

class FakeXhr {
  open() {}
  setRequestHeader() {}
  send() {}
  abort() {}
  addEventListener() {}
  removeEventListener() {}
}

function loadCompassModule() {
  const source = fs.readFileSync(userscriptPath, 'utf8');
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
