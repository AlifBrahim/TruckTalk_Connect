/**
 * Lightweight unit tests for pure utilities and mapping.
 * Run via Apps Script editor by executing runUnitTests().
 */

var __TESTS__ = [];

function test(name, fn) { __TESTS__.push({ name: name, fn: fn }); }

function assertTrue(cond, msg) {
  if (!cond) throw new Error('Assertion failed: ' + (msg || '')); 
}

function assertEquals(actual, expected, msg) {
  var a = (typeof actual === 'object') ? JSON.stringify(actual) : String(actual);
  var e = (typeof expected === 'object') ? JSON.stringify(expected) : String(expected);
  if (a !== e) throw new Error('Assertion failed: ' + (msg || '') + ' Expected ' + e + ', got ' + a);
}

// ---- Tests ---------------------------------------------------------------

test('isISO8601_ basic', function () {
  assertTrue(isISO8601_('2025-08-29T20:00:00Z'), 'valid ISO should be true');
  assertTrue(!isISO8601_('2025-08-29 20:00:00Z'), 'space instead of T should be false');
  assertTrue(!isISO8601_('2025-08-29T20:00Z'), 'missing seconds should be false');
});

test('stringHasExplicitTZ_ detects explicit tz', function () {
  assertTrue(stringHasExplicitTZ_('2025-08-29T20:00:00Z'), 'Z suffix');
  assertTrue(stringHasExplicitTZ_('2025-08-29 14:00 -0600'), 'offset -0600');
  assertTrue(stringHasExplicitTZ_('2025-08-29 14:00 -06:00'), 'offset -06:00');
  assertTrue(stringHasExplicitTZ_('Fri Aug 29 2025 14:00:00 GMT-0600 (CST)'), 'GMT offset');
  assertTrue(!stringHasExplicitTZ_('2025-08-29 14:00'), 'no tz');
});

test('valueHasExplicitTZ_ string only', function () {
  assertTrue(valueHasExplicitTZ_('2025-08-29T20:00:00Z'), 'Z');
  assertTrue(!valueHasExplicitTZ_(new Date()), 'Date object returns false');
  assertTrue(!valueHasExplicitTZ_(45200), 'numeric serial returns false');
});

test('canon_ and findHeaderBySynonyms_', function () {
  var headers = ['PU Time', 'DEL Time', 'Pickup Address'];
  var hit1 = findHeaderBySynonyms_(headers, ['pu time']);
  var hit2 = findHeaderBySynonyms_(headers, ['delivery time','del time']);
  assertEquals(hit1, 'PU Time');
  assertEquals(hit2, 'DEL Time');
});

test('extractTime_ parses strings and numbers', function () {
  var t1 = extractTime_('2:30 PM');
  assertEquals(t1 && t1.h, 14, '2:30 PM hour');
  assertEquals(t1 && t1.m, 30, '2:30 PM minute');

  var t2 = extractTime_(0.5); // Excel half-day → 12:00:00
  assertEquals(t2 && t2.h, 12, '0.5 hour');
  assertEquals(t2 && t2.m, 0, '0.5 minute');
});

test('extractYMD_ parses strings', function () {
  var ymd1 = extractYMD_('2025-08-29');
  assertEquals(ymd1 && ymd1.y, 2025, 'ISO year');
  assertEquals(ymd1 && ymd1.m, 7, 'ISO month zero-based for Aug');
  assertEquals(ymd1 && ymd1.d, 29, 'ISO day');

  var ymd2 = extractYMD_('8/29/2025');
  assertEquals(ymd2 && ymd2.y, 2025, 'US year');
  assertEquals(ymd2 && ymd2.m, 7, 'US month zero-based for Aug');
  assertEquals(ymd2 && ymd2.d, 29, 'US day');
});

test('combineLocalDateTimeToUTC_ requires both parts', function () {
  var both = combineLocalDateTimeToUTC_('2025-08-29', '14:00');
  assertTrue(!!both, 'returns ISO-like string when both provided');
  var missingDate = combineLocalDateTimeToUTC_(null, '14:00');
  var missingTime = combineLocalDateTimeToUTC_('2025-08-29', null);
  assertTrue(missingDate === null, 'null when date missing');
  assertTrue(missingTime === null, 'null when time missing');
});

test('headerIndex_ builds index map', function () {
  var idx = headerIndex_(['A','B','C']);
  assertEquals(idx.A, 0);
  assertEquals(idx.C, 2);
});

test('strictResult_ only includes loads when ok===true', function () {
  var r1 = strictResult_({ ok: false, issues: [], mapping: {}, meta: { analyzedRows: 0, analyzedAt: 'x' }, loads: [{a:1}] });
  assertTrue(!('loads' in r1), 'loads omitted on ok=false');
  var r2 = strictResult_({ ok: true, issues: [], mapping: {}, meta: { analyzedRows: 0, analyzedAt: 'x' }, loads: [{a:1}] });
  assertTrue(Array.isArray(r2.loads), 'loads present on ok=true');
});

test('issue_ shapes fields correctly', function () {
  var it = issue_('CODE','warn','msg',[2], 'col', 'suggest');
  assertEquals(it.code, 'CODE');
  assertEquals(it.severity, 'warn');
  assertEquals(it.message, 'msg');
  assertEquals(JSON.stringify(it.rows), JSON.stringify([2]));
  assertEquals(it.column, 'col');
  assertEquals(it.suggestion, 'suggest');
});

test('detectHeaderMapping_ maps VRID → loadId', function () {
  var res = detectHeaderMapping_(['VRID','PU','DEL'], {});
  assertEquals(res.mapping['VRID'], 'loadId');
});

/**
 * Run all tests and return a summary object.
 */
function runUnitTests() {
  var passed = 0, failed = 0, failures = [];
  for (var i = 0; i < __TESTS__.length; i++) {
    var t = __TESTS__[i];
    try {
      t.fn();
      passed++;
      Logger.log('PASS: ' + t.name);
    } catch (e) {
      failed++;
      var msg = 'FAIL: ' + t.name + ' → ' + (e && e.message || e);
      failures.push(msg);
      Logger.log(msg);
    }
  }
  var summary = { total: __TESTS__.length, passed: passed, failed: failed, failures: failures };
  Logger.log('Summary: ' + JSON.stringify(summary));
  return summary;
}

