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
 *   ?year=2027 → 공통 탭 + 2027_ 탭만 보냄 (다른 학년 탭은 보내지 않아 빠름)
 *   ?list=1    → 기본정보 탭만 보냄 (입학년도 선택 화면용)
 *
 * 탭을 추가하거나 이름을 바꿔도 이 코드는 고칠 필요가 없습니다.
 * 코드를 고친 경우에만 다시 배포: [배포] → [배포 관리] → 연필(수정) → 버전: [새 버전] → [배포]
 *   (※ [새 배포]를 누르면 주소가 바뀌니 주의)
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var year = String(params.year || '');
  var listOnly = params.list === '1';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var finalData = {};

  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name.charAt(0) === '_') return;
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

  return ContentService.createTextOutput(JSON.stringify(finalData))
    .setMimeType(ContentService.MimeType.JSON);
}
