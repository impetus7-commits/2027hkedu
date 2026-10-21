/**
 * 과목 선택 시스템 — 시트 → 웹페이지 데이터 변환 (Apps Script)
 *
 * 시트 하나로 3개 학년을 관리합니다.
 *   - 공통 탭: 이름에 입학년도가 없는 탭 (예: 기본정보, 계열데이터, 학과데이터)
 *   - 학년별 탭: 이름이 "입학년도_"로 시작하는 탭 (예: 2027_과목데이터, 2027_규칙)
 *   - 이름이 _ 로 시작하는 탭은 보내지 않습니다. (예: _사용안내, _학생명단)
 *     ⚠️ 학생 개인정보가 있는 탭은 반드시 이름 앞에 _ 를 붙이세요.
 *
 * 웹페이지가 부르는 방식
 *   ?year=2027        → 공통 탭 + 2027_ 탭만 보냄 (다른 학년 탭은 보내지 않아 빠름)
 *   ?list=1           → 기본정보 탭만 보냄 (입학년도 선택 화면용)
 *   ?univ=간호학과     → [대학학과정보] 탭에서 그 학과를 개설한 대학만 보냄 (관련 학과·대학 살펴보기)
 *   ?cjob=간호사       → [커리어넷직업] 탭에서 그 직업의 직업백과 내용만 보냄
 *   ?counsel=간호사    → 커리어넷 진로상담 사례를 그때그때 찾아서 보냄 (6시간 동안 저장해 두고 재사용)
 *
 * [대학학과정보]·[커리어넷직업] 탭은 시트 위쪽 메뉴 [🎓 대학 정보]로 채웁니다. (1년에 한 번)
 *   처음 한 번은 [🔑 커리어넷 인증키 저장]으로 인증키를 저장하세요. 키는 이 시트의 스크립트에만 보관되고 학생에게 보이지 않습니다.
 *
 * 탭을 추가하거나 이름을 바꿔도 이 코드는 고칠 필요가 없습니다.
 * 코드를 고친 경우에만 다시 배포: [배포] → [배포 관리] → 연필(수정) → 버전: [새 버전] → [배포]
 *   (※ [새 배포]를 누르면 주소가 바뀌니 주의)
 */
var UNIV_TAB = '대학학과정보';            // 양이 많아 ?year= 요청에는 넣지 않고, ?univ= 요청 때만 찾아서 보냄
var JOB_TAB = '커리어넷직업';            // 직업백과 (양이 많아 ?cjob= 요청 때만 찾아서 보냄)
var ON_DEMAND_TABS = [UNIV_TAB, JOB_TAB];

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.univ) return json_(findUniversities_(String(params.univ)));
  if (params.cjob) return json_(findCareerJobs_(String(params.cjob)));
  if (params.counsel) return json_(findCounsel_(String(params.counsel)));

  var year = String(params.year || '');
  var listOnly = params.list === '1';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var finalData = {};

  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name.charAt(0) === '_') return;
    if (ON_DEMAND_TABS.indexOf(name.replace(/\s+/g, '')) !== -1) return;
    if (listOnly && name.replace(/\s+/g, '') !== '기본정보') return;

    var m = name.match(/^(\d{4})\s*[_\-]/);
    if (m && m[1] !== year) return;   // 다른 입학년도 탭은 건너뜀

    var rows = sheet.getDataRange().getDisplayValues()
      .filter(function (row) { return row.some(function (cell) { return cell !== ''; }); });
    if (rows.length === 0) return;

    // 오른쪽 끝의 빈 열 잘라내기 (전송량 줄이기)
    var width = 0;
    rows.forEach(function (row) {
      for (var i = row.length - 1; i >= width; i--) { if (row[i] !== '') { width = i + 1; break; } }
    });
    finalData[name] = rows.map(function (row) { return row.slice(0, width); });
  });

  return json_(finalData);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────── 관련 학과 → 개설 대학 찾기 (?univ=) ─────────────
// 학과 이름에서 '학과·학부·전공·과'를 떼고 비교 (간호학과 = 간호학부 = 간호학전공 = 간호과)
function majorCore_(s) {
  return String(s || '').replace(/\(.*?\)/g, '').replace(/[^0-9A-Za-z가-힣]/g, '')
    .replace(/(학과|학부|전공|계열|과)$/, '');
}

function findUniversities_(query) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'univ:' + (PropertiesService.getScriptProperties().getProperty('UNIV_VERSION') || '') + ':' + query;
  var hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIV_TAB);
  if (!sheet || sheet.getLastRow() < 2) return { ready: false, rows: [] };
  var values = sheet.getDataRange().getDisplayValues();
  var head = values[0];
  var col = function (name) { return head.indexOf(name); };
  var c = { std: col('표준학과'), field: col('계열'), kind: col('구분'), univ: col('대학명'), campus: col('캠퍼스'),
            area: col('소재지'), dept: col('대학 학과명'), url: col('홈페이지') };

  var cores = query.split('|').map(majorCore_).filter(function (q) { return q.length >= 2; });
  var rows = [];
  for (var i = 1; i < values.length && rows.length < 600; i++) {
    var r = values[i];
    var std = majorCore_(r[c.std]), dept = majorCore_(r[c.dept]);
    var matched = cores.some(function (q) { return std === q || dept === q || (q.length >= 3 && dept.indexOf(q) !== -1); });
    if (matched) rows.push([r[c.std], r[c.field], r[c.kind], r[c.univ], r[c.campus], r[c.area], r[c.dept], r[c.url]]);
  }
  var result = { ready: true, updated: String(sheet.getRange(1, 10).getDisplayValue() || ''),
                 head: ['표준학과', '계열', '구분', '대학명', '캠퍼스', '소재지', '대학 학과명', '홈페이지'], rows: rows };
  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (err) { /* 너무 크면 캐시 생략 */ }
  return result;
}

// ───────────── 시트 메뉴: 커리어넷에서 [대학학과정보] 탭 채우기 ─────────────
var CAREERNET_API = 'https://www.career.go.kr/cnet/openapi/getOpenApi';
var RUN_LIMIT_MS = 4.5 * 60 * 1000;      // Apps Script 한 번 실행 한도(6분) 안에서 끊고 이어서 실행

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎓 대학 정보')
    .addItem('🔑 커리어넷 인증키 저장', 'saveCareerNetKey')
    .addItem('⬇️ 커리어넷에서 불러오기 (1년에 한 번)', 'startUniversityImport')
    .addItem('📘 커리어넷 직업백과 불러오기 (1년에 한 번)', 'startJobImport')
    .addItem('⏹ 불러오기 멈추기', 'stopUniversityImport')
    .addToUi();
}

function saveCareerNetKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('커리어넷 오픈API 인증키', '발급받은 인증키를 붙여 넣으세요. (학생에게는 보이지 않습니다)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty('CAREERNET_KEY', res.getResponseText().trim());
  ui.alert('저장했습니다. 이제 [🎓 대학 정보] → [커리어넷에서 불러오기]를 누르세요.');
}

function careerNet_(params) {
  var key = PropertiesService.getScriptProperties().getProperty('CAREERNET_KEY');
  if (!key) throw new Error('커리어넷 인증키가 없습니다. [🎓 대학 정보] → [🔑 커리어넷 인증키 저장]을 먼저 하세요.');
  var q = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  return CAREERNET_API + '?apiKey=' + encodeURIComponent(key) + '&svcType=api&contentType=json&' + q;
}

function contentOf_(text) {
  var data = JSON.parse(text);
  var c = data && data.dataSearch && data.dataSearch.content;
  return Array.isArray(c) ? c : (c ? [c] : []);
}

// 학교 이름으로 4년제/전문대 구분 (커리어넷 자료에 구분이 없을 때)
function schoolKind_(name) {
  name = String(name || '');
  if (/대학교|과학기술원|KAIST|DGIST|UNIST|GIST|한국예술종합학교|사관학교|교육대학/.test(name)) return '4년제';
  if (/전문대학|대학$|폴리텍/.test(name)) return '전문대';
  return '기타';
}

function startUniversityImport() {
  var props = PropertiesService.getScriptProperties();
  // 1) 표준학과 목록 받기
  var list = [];
  ['univ_list'].forEach(function (gubun) {
    var res = UrlFetchApp.fetch(careerNet_({ svcCode: 'MAJOR', gubun: gubun, perPage: 2000, thisPage: 1 }), { muteHttpExceptions: true });
    saveSample_('목록', res.getContentText());
    contentOf_(res.getContentText()).forEach(function (m) {
      if (m.majorSeq) list.push({ seq: String(m.majorSeq), name: m.mClass || m.majorName || '', field: m.lClass || '' });
    });
  });
  if (list.length === 0) throw new Error('커리어넷에서 학과 목록을 받지 못했습니다. [_커리어넷_응답예시] 탭을 확인하세요.');

  // 2) 탭 새로 만들기
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(UNIV_TAB) || ss.insertSheet(UNIV_TAB);
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).setValues([['표준학과', '계열', '구분', '대학명', '캠퍼스', '소재지', '대학 학과명', '홈페이지']])
    .setFontWeight('bold').setBackground('#065f46').setFontColor('#ffffff');
  sheet.getRange(1, 10).setValue('불러오는 중…');
  sheet.setFrozenRows(1);
  props.setProperty('UNIV_IMPORT', JSON.stringify({ list: list, next: 0 }));
  continueUniversityImport();
}

function continueUniversityImport() {
  var started = Date.now();
  var props = PropertiesService.getScriptProperties();
  var state = JSON.parse(props.getProperty('UNIV_IMPORT') || 'null');
  clearTriggers_('continueUniversityImport');
  if (!state) return;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIV_TAB);

  while (state.next < state.list.length) {
    if (Date.now() - started > RUN_LIMIT_MS) {   // 시간이 다 되면 1분 뒤 이어서
      props.setProperty('UNIV_IMPORT', JSON.stringify(state));
      ScriptApp.newTrigger('continueUniversityImport').timeBased().after(60 * 1000).create();
      sheet.getRange(1, 10).setValue('불러오는 중… ' + state.next + '/' + state.list.length + ' 학과 (자동으로 이어서 진행)');
      return;
    }
    var chunk = state.list.slice(state.next, state.next + 20);
    var responses = UrlFetchApp.fetchAll(chunk.map(function (m) {
      return { url: careerNet_({ svcCode: 'MAJOR_VIEW', gubun: 'univ_list', majorSeq: m.seq }), muteHttpExceptions: true };
    }));
    var out = [];
    responses.forEach(function (res, i) {
      var m = chunk[i];
      if (state.next === 0 && i === 0) saveSample_('학과상세', res.getContentText());
      var content;
      try { content = contentOf_(res.getContentText()); } catch (err) { return; }
      content.forEach(function (d) {
        var univs = d.university || d.univ || [];
        if (!Array.isArray(univs)) univs = [univs];
        univs.forEach(function (u) {
          var school = u.schoolName || u.univ_nm || '';
          if (!school) return;
          out.push([m.name, m.field, u.schoolGubun || schoolKind_(school), school, u.campus_nm || u.campusName || '',
                    u.area || u.region || '', u.majorName || u.major_nm || '', u.schoolURL || u.url || '']);
        });
      });
    });
    if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, 8).setValues(out);
    state.next += chunk.length;
  }
  props.deleteProperty('UNIV_IMPORT');
  sheet.getRange(1, 10).setValue('불러온 날: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd') + ' (커리어넷 학과정보)');
  props.setProperty('UNIV_VERSION', String(Date.now()));   // 예전 검색 결과 캐시를 쓰지 않도록
  try { SpreadsheetApp.getActiveSpreadsheet().toast('대학 정보 불러오기를 마쳤습니다. (' + (sheet.getLastRow() - 1) + '줄)', '🎓 대학 정보', 10); } catch (err) {}
}

function stopUniversityImport() {
  clearTriggers_();
  PropertiesService.getScriptProperties().deleteProperty('UNIV_IMPORT');
  PropertiesService.getScriptProperties().deleteProperty('JOB_IMPORT');
  SpreadsheetApp.getActiveSpreadsheet().toast('멈췄습니다. 다시 하려면 [커리어넷에서 불러오기]를 누르세요.');
}

function clearTriggers_(only) {
  var names = only ? [only] : ['continueUniversityImport', 'continueJobImport'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}

// 커리어넷 응답 모양이 예상과 다를 때 확인용 (앞부분만 저장, 웹페이지로는 보내지 않음)
function saveSample_(label, text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_커리어넷_응답예시') || ss.insertSheet('_커리어넷_응답예시');
  if (sh.getLastRow() > 40) return;   // 너무 쌓이지 않게
  sh.appendRow([new Date(), label, String(text).slice(0, 45000)]);
}

// ───────────── 커리어넷 직업백과 → [커리어넷직업] 탭 ─────────────
//  커리어넷 응답 모양이 두 가지(새 직업백과 API / 예전 getOpenApi)라 둘 다 받아 들임
var JOB_HEAD = ['직업명', '하는 일', '적성·흥미', '관련 학과', '관련 자격', '임금', '전망', '되는 길', '직업코드', '원문'];
var CAREERNET_JOBS = 'https://www.career.go.kr/cnet/front/openapi/jobs.json';
var CAREERNET_JOB = 'https://www.career.go.kr/cnet/front/openapi/job.json';

function careerKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('CAREERNET_KEY');
  if (!key) throw new Error('커리어넷 인증키가 없습니다. [🎓 대학 정보] → [🔑 커리어넷 인증키 저장]을 먼저 하세요.');
  return key;
}
function tryJson_(text) { try { return JSON.parse(text); } catch (e) { return null; } }

// 객체 안(두 단계까지)에서 이름이 맞는 값 찾기
function findKey_(o, re) {
  if (!o || typeof o !== 'object') return null;
  var keys = Object.keys(o), i, v;
  for (i = 0; i < keys.length; i++) if (re.test(keys[i]) && o[keys[i]] !== '' && o[keys[i]] != null) return o[keys[i]];
  for (i = 0; i < keys.length; i++) { v = o[keys[i]]; if (v && typeof v === 'object' && !Array.isArray(v)) { var r = findKey_(v, re); if (r != null) return r; } }
  return null;
}
// 값(문자열·배열·객체)을 읽기 좋은 글로
function textOf_(v, sep) {
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number') return String(v).replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
  if (Array.isArray(v)) return v.map(function (x) {
    if (x && typeof x === 'object') {
      var pref = findKey_(x, /^(work|name|.*_name|.*_nm|certi|depart_name|forecast|prepare.*|content|title|.*Name)$/i);
      if (pref != null && typeof pref !== 'object') return textOf_(pref);
      return Object.keys(x).map(function (k) { return typeof x[k] === 'string' ? x[k] : ''; }).filter(String).join(' ');
    }
    return textOf_(x);
  }).filter(String).join(sep || '\n');
  if (typeof v === 'object') return Object.keys(v).map(function (k) { return typeof v[k] === 'string' ? v[k] : ''; }).filter(String).join(' ');
  return '';
}
function jobRow_(d, fallbackCode) {
  var code = String(findKey_(d, /^(job_cd|jobdicSeq|seq|job_code)$/i) || fallbackCode || '');
  var name = textOf_(findKey_(d, /^(job_nm|job|jobName|job_name)$/i));
  var interest = textOf_(findKey_(d, /^(interestList|interest)$/i), ', ');
  var apt = textOf_(findKey_(d, /^(aptitudeList|abilityList|aptitude|aptit_name)$/i), ', ');
  return [name,
    textOf_(findKey_(d, /^(workList|work|summary|job_summary)$/i)),
    [apt, interest].filter(String).join(' / '),
    textOf_(findKey_(d, /^(departList|depart|relDepartList|major)$/i), ', '),
    textOf_(findKey_(d, /^(certiList|certi|certificate)$/i), ', '),
    textOf_(findKey_(d, /^(wage|salery|salary)$/i)),
    textOf_(findKey_(d, /^(forecastList|forecast|prospect|possibility)$/i)),
    textOf_(findKey_(d, /^(prepareList|preparation|prepare|jobRoad|way)$/i)),
    code,
    code ? 'https://www.career.go.kr/cnet/front/base/job/jobView.do?SEQ=' + code : ''];
}

function startJobImport() {
  var key = careerKey_();
  var props = PropertiesService.getScriptProperties();
  var list = [], legacyRows = [];
  // 1) 새 직업백과 API: 쪽마다 목록
  for (var page = 1; page <= 30; page++) {
    var res = UrlFetchApp.fetch(CAREERNET_JOBS + '?apiKey=' + encodeURIComponent(key) + '&pageIndex=' + page, { muteHttpExceptions: true });
    if (page === 1) saveSample_('직업 목록(새)', res.getContentText());
    var data = tryJson_(res.getContentText());
    var jobs = data && (data.jobs || findKey_(data, /^jobs$/));
    if (!Array.isArray(jobs) || jobs.length === 0) break;
    jobs.forEach(function (j) { var c = findKey_(j, /^(job_cd|seq)$/i); if (c) list.push(String(c)); });
  }
  // 2) 안 되면 예전 getOpenApi (목록에 설명이 같이 들어 있음)
  if (list.length === 0) {
    var r2 = UrlFetchApp.fetch(careerNet_({ svcCode: 'JOB', gubun: 'job_dic_list', perPage: 2000, thisPage: 1 }), { muteHttpExceptions: true });
    saveSample_('직업 목록(예전)', r2.getContentText());
    contentOf_(r2.getContentText()).forEach(function (d) { var row = jobRow_(d); if (row[0]) legacyRows.push(row); });
    if (legacyRows.length === 0) throw new Error('커리어넷에서 직업 목록을 받지 못했습니다. [_커리어넷_응답예시] 탭을 확인하세요.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(JOB_TAB) || ss.insertSheet(JOB_TAB);
  sheet.clear();
  sheet.getRange(1, 1, 1, JOB_HEAD.length).setValues([JOB_HEAD]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  if (legacyRows.length) {
    sheet.getRange(2, 1, legacyRows.length, JOB_HEAD.length).setValues(legacyRows);
    finishJobImport_(sheet);
    return;
  }
  sheet.getRange(1, 12).setValue('불러오는 중…');
  props.setProperty('JOB_IMPORT', JSON.stringify({ list: list, next: 0 }));
  continueJobImport();
}

function continueJobImport() {
  var started = Date.now();
  var props = PropertiesService.getScriptProperties();
  var state = JSON.parse(props.getProperty('JOB_IMPORT') || 'null');
  clearTriggers_('continueJobImport');
  if (!state) return;
  var key = careerKey_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOB_TAB);
  while (state.next < state.list.length) {
    if (Date.now() - started > RUN_LIMIT_MS) {
      props.setProperty('JOB_IMPORT', JSON.stringify(state));
      ScriptApp.newTrigger('continueJobImport').timeBased().after(60 * 1000).create();
      sheet.getRange(1, 12).setValue('불러오는 중… ' + state.next + '/' + state.list.length + ' 직업 (자동으로 이어서 진행)');
      return;
    }
    var chunk = state.list.slice(state.next, state.next + 20);
    var responses = UrlFetchApp.fetchAll(chunk.map(function (c) {
      return { url: CAREERNET_JOB + '?apiKey=' + encodeURIComponent(key) + '&seq=' + encodeURIComponent(c), muteHttpExceptions: true };
    }));
    var out = [];
    responses.forEach(function (res, i) {
      if (state.next === 0 && i === 0) saveSample_('직업 상세(새)', res.getContentText());
      var d = tryJson_(res.getContentText());
      if (!d) return;
      var row = jobRow_(d, chunk[i]);
      if (row[0]) out.push(row);
    });
    if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, JOB_HEAD.length).setValues(out);
    state.next += chunk.length;
  }
  props.deleteProperty('JOB_IMPORT');
  finishJobImport_(sheet);
}

function finishJobImport_(sheet) {
  sheet.getRange(1, 12).setValue('불러온 날: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd') + ' (커리어넷 직업백과)');
  PropertiesService.getScriptProperties().setProperty('JOB_VERSION', String(Date.now()));
  try { SpreadsheetApp.getActiveSpreadsheet().toast('직업백과 불러오기를 마쳤습니다. (' + (sheet.getLastRow() - 1) + '개 직업)', '🎓 대학 정보', 10); } catch (err) {}
}

// ?cjob=간호사|보건교사 → 이름이 같거나 서로 포함되는 직업 (최대 4개)
function jobCore_(s) { return String(s || '').replace(/\(.*?\)/g, '').replace(/[^0-9A-Za-z가-힣]/g, ''); }
// '항공승무원' ↔ '항공기객실승무원': 첫 글자·끝 두 글자가 같고, 나머지 글자가 순서대로 다 들어 있으면 비슷한 직업
function looseJob_(q, n) {
  if (q.length < 3 || n.length > q.length + 6 || q.charAt(0) !== n.charAt(0) || q.slice(-2) !== n.slice(-2)) return false;
  for (var i = 0, j = 0; j < n.length && i < q.length; j++) if (q.charAt(i) === n.charAt(j)) i++;
  return i === q.length;
}
function findCareerJobs_(query) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'cjob:' + (PropertiesService.getScriptProperties().getProperty('JOB_VERSION') || '') + ':' + query;
  var hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOB_TAB);
  if (!sheet || sheet.getLastRow() < 2) return { ready: false, rows: [] };
  var values = sheet.getDataRange().getDisplayValues();
  var qs = query.split('|').map(jobCore_).filter(function (q) { return q.length >= 2; });
  var exact = [], near = [];
  for (var i = 1; i < values.length; i++) {
    var n = jobCore_(values[i][0]);
    if (!n) continue;
    if (qs.indexOf(n) !== -1) exact.push(values[i].slice(0, JOB_HEAD.length));
    else if (qs.some(function (q) { return (q.length >= 3 && n.indexOf(q) !== -1) || (n.length >= 3 && q.indexOf(n) !== -1) || looseJob_(q, n); })) near.push(values[i].slice(0, JOB_HEAD.length));
  }
  var result = { ready: true, updated: String(sheet.getRange(1, 12).getDisplayValue() || ''), head: JOB_HEAD, rows: exact.concat(near).slice(0, 4), exact: exact.length };
  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (err) {}
  return result;
}

// ───────────── 커리어넷 진로상담 사례 (?counsel=낱말) ─────────────
//  그때그때 커리어넷에 물어보고 6시간 저장. 목록과 앞부분만 보내고, 전문은 커리어넷 원문으로 연결
function findCounsel_(word) {
  word = String(word).slice(0, 30);
  var cache = CacheService.getScriptCache();
  var cacheKey = 'counsel:' + word;
  var hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);
  var key = PropertiesService.getScriptProperties().getProperty('CAREERNET_KEY');
  if (!key) return { ready: false, items: [] };
  var tries = [
    { svcCode: 'COUNSEL', gubun: 'counsel_list', searchTitle: word, perPage: 8, thisPage: 1 },
    { svcCode: 'COUNSEL', searchTitle: word, perPage: 8, thisPage: 1 },
    { svcCode: 'COUNSEL', searchKeyword: word, perPage: 8, thisPage: 1 }
  ];
  var items = [];
  for (var t = 0; t < tries.length && items.length === 0; t++) {
    var res = UrlFetchApp.fetch(careerNet_(tries[t]), { muteHttpExceptions: true });
    saveSample_('상담 ' + t, res.getContentText());
    var content = [];
    try { content = contentOf_(res.getContentText()); } catch (err) { content = []; }
    content.forEach(function (c) {
      var title = textOf_(findKey_(c, /^(title|subject|qTitle|question_title)$/i));
      if (!title) return;
      var link = textOf_(findKey_(c, /^(url|link|linkUrl|detailUrl)$/i));
      var seq = textOf_(findKey_(c, /^(seq|counselSeq|idx|no)$/i));
      items.push({
        title: title.slice(0, 120),
        q: textOf_(findKey_(c, /^(question|qContent|content|qcontent)$/i)).slice(0, 400),
        a: textOf_(findKey_(c, /^(answer|aContent|acontent|reply)$/i)).slice(0, 600),
        field: textOf_(findKey_(c, /^(category|field|gubun|counsel_type)$/i)).slice(0, 30),
        link: link || ''
      });
    });
  }
  var result = { ready: true, items: items.slice(0, 6) };
  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (err) {}
  return result;
}
