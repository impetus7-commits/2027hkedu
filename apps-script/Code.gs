/**
 * 과목 선택 시스템 — 시트 → 웹페이지 데이터 변환 (Apps Script)
 *
 * - 시트의 모든 탭을 자동으로 보냅니다. 탭을 추가하거나 이름을 바꿔도 이 코드는 고칠 필요가 없습니다.
 * - 이름이 _ 로 시작하는 탭은 보내지 않습니다. (예: _학생명단, _작성법, _메모)
 *   ⚠️ 학생 개인정보가 있는 탭은 반드시 이름 앞에 _ 를 붙이세요.
 * - 열 순서를 그대로 유지하고, 시트에 보이는 그대로(글자) 보냅니다.
 *
 * 수정 후 배포: [배포] → [배포 관리] → 연필(수정) → 버전: [새 버전] → [배포]
 *   (※ [새 배포]를 누르면 주소가 바뀌니 주의)
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalData = {};

  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    if (name.charAt(0) === '_') return;

    const rows = sheet.getDataRange().getDisplayValues()
      .filter(function (row) { return row.some(function (cell) { return cell !== ''; }); });
    if (rows.length > 0) finalData[name] = rows;
  });

  return ContentService.createTextOutput(JSON.stringify(finalData))
    .setMimeType(ContentService.MimeType.JSON);
}
