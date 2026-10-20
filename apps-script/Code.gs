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
 *
 * [대학학과정보] 탭은 시트 위쪽 메뉴 [🎓 대학 정보] → [커리어넷에서 불러오기]로 채웁니다. (1년에 한 번)
 *   처음 한 번은 [🔑 커리어넷 인증키 저장]으로 인증키를 저장하세요. 키는 이 시트의 스크립트에만 보관되고 학생에게 보이지 않습니다.
 *
 * 탭을 추가하거나 이름을 바꿔도 이 코드는 고칠 필요가 없습니다.
 * 코드를 고친 경우에만 다시 배포: [배포] → [배포 관리] → 연필(수정) → 버전: [새 버전] → [배포]
 *   (※ [새 배포]를 누르면 주소가 바뀌니 주의)
 */
var UNIV_TAB = '대학학과정보';            // 양이 많아 ?year= 요청에는 넣지 않고, ?univ= 요청 때만 찾아서 보냄
var ON_DEMAND_TABS = [UNIV_TAB];

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.univ) return json_(findUniversities_(String(params.univ)));

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
  clearTriggers_();
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
  SpreadsheetApp.getActiveSpreadsheet().toast('멈췄습니다. 다시 하려면 [커리어넷에서 불러오기]를 누르세요.');
}

function clearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continueUniversityImport') ScriptApp.deleteTrigger(t);
  });
}

// 커리어넷 응답 모양이 예상과 다를 때 확인용 (앞부분만 저장, 웹페이지로는 보내지 않음)
function saveSample_(label, text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_커리어넷_응답예시') || ss.insertSheet('_커리어넷_응답예시');
  sh.appendRow([new Date(), label, String(text).slice(0, 45000)]);
}
