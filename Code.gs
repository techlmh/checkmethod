function doGet(e) {
  // ?mode=admin 이면 관리자 페이지
  if (e.parameter && e.parameter.mode === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('관리자용 제출 내역 확인')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // ?mode=school 이면 학교 담당자용 페이지
  if (e.parameter && e.parameter.mode === 'school') {
    return HtmlService.createHtmlOutputFromFile('School')
      .setTitle('학교 담당자 컨설팅 반영 시스템')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // 기본은 입력(위원용) 페이지
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('중학교 교육과정 컨설팅 입력 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function extractFileId(url) {
  if (!url) return null;
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// [공통] 시트를 안전하게 가져오기 (첫번째 시트가 메인 데이터라고 가정)
function getMainSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// -------------------------------------------------------------
// [위원용] 데이터 저장 (기존과 동일)
// -------------------------------------------------------------
function submitData(data) {
  try {
    var sheet = getMainSheet();
    var folderId = "1EbWiQrkiHadHAm_aw0vZY5i8xmAndZrD"; // 선생님 폴더 ID 유지
    var folder = DriveApp.getFolderById(folderId);
    var newImageUrls = data.items.map(function(item) { return item.imageUrl; });

    if (data.originalTimestamp) {
      var sheetData = sheet.getDataRange().getDisplayValues();
      for (var i = sheetData.length - 1; i >= 1; i--) {
        if (sheetData[i][0] === data.originalTimestamp && sheetData[i][1] === data.consultantName && sheetData[i][2] === data.schoolName) {
          var oldUrl = sheetData[i][6]; 
          if (oldUrl && newImageUrls.indexOf(oldUrl) === -1) {
            var fileId = extractFileId(oldUrl);
            if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
          }
          sheet.deleteRow(i + 1);
        }
      }
    }

    var timestamp = new Date();
    var dateString = Utilities.formatDate(timestamp, "Asia/Seoul", "yyyyMMdd_HHmmss");
    var formattedTime = Utilities.formatDate(timestamp, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    data.items.forEach(function(item) {
      var finalImageUrl = item.imageUrl || ""; 
      if (item.imageBase64) {
        var base64Data = item.imageBase64.split(',')[1];
        var safeSchoolName = data.schoolName.replace(/[/\\?%*:|"<>]/g, ''); 
        var fileName = safeSchoolName + "_" + item.grade + "_" + item.type + "_" + dateString + ".png";
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), item.imageType || 'image/png', fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        finalImageUrl = file.getUrl();
      }
      sheet.appendRow([formattedTime, data.consultantName, data.schoolName, item.grade, item.type, item.content, finalImageUrl]);
    });
    return { success: true };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function getServerHistory(searchConsultant, searchSchool) {
  var sheet = getMainSheet();
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  var historyMap = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== searchConsultant || data[i][2] !== searchSchool) continue; 
    var ts = data[i][0];
    if (!historyMap[ts]) { historyMap[ts] = { timestamp: ts, consultantName: data[i][1], schoolName: data[i][2], items: [] }; }
    historyMap[ts].items.push({ grade: data[i][3], type: data[i][4], content: data[i][5], imageUrl: data[i][6], imageBase64: '' });
  }
  var result = Object.keys(historyMap).map(function(k) { return historyMap[k]; });
  result.reverse(); return result;
}

// -------------------------------------------------------------
// [관리자용] 데이터 불러오기
// -------------------------------------------------------------
function getAllDataForAdmin() {
  var sheet = getMainSheet();
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] + "_" + data[i][1] + "_" + data[i][2];
    if (!groups[key]) {
      groups[key] = { id: key, timestamp: data[i][0], consultantName: data[i][1], schoolName: data[i][2], items: [] };
    }
    groups[key].items.push({ grade: data[i][3], type: data[i][4], content: data[i][5], imageUrl: data[i][6] });
  }
  var result = Object.keys(groups).map(function(k) { return groups[k]; });
  result.reverse(); return result;
}

function getImagesBase64Mapping(urls) {
  var map = {};
  for(var i = 0; i < urls.length; i++) {
    var url = urls[i];
    if(!url || map[url]) continue;
    var fileId = extractFileId(url);
    if(fileId) {
      try {
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        map[url] = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
      } catch(e) { map[url] = ''; }
    }
  }
  return map;
}

// -------------------------------------------------------------
// [학교 담당자용] 나이스 번호로 데이터 조회 및 반영여부 저장
// -------------------------------------------------------------
function getSchoolDataByNeis(neisCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mappingSheet = ss.getSheetByName('학교코드');
  if (!mappingSheet) throw new Error("'학교코드' 시트가 존재하지 않습니다. 먼저 시트를 만들어주세요.");
  
  var mappingData = mappingSheet.getDataRange().getDisplayValues();
  var schoolName = "";
  
  for(var i=1; i<mappingData.length; i++) {
    if(mappingData[i][0].replace(/\s/g, '') === neisCode.replace(/\s/g, '')) {
      schoolName = mappingData[i][1];
      break;
    }
  }
  
  if(!schoolName) throw new Error("일치하는 나이스 번호가 없습니다.");
  
  var mainSheet = getMainSheet();
  var mainData = mainSheet.getDataRange().getDisplayValues();
  var items = [];
  
  // 담당자 이름(H열=7), 반영여부(I열=8) 가져오기
  var repName = ""; 
  for(var i=1; i<mainData.length; i++) {
    if(mainData[i][2] === schoolName) {
      items.push({
        rowIndex: i + 1, // 수정 시 정확한 행 위치 기록
        grade: mainData[i][3],
        type: mainData[i][4],
        content: mainData[i][5],
        imageUrl: mainData[i][6],
        reflection: mainData[i][8] || ""
      });
      if(mainData[i][7]) repName = mainData[i][7];
    }
  }
  
  return { schoolName: schoolName, repName: repName, items: items };
}

function submitSchoolReflection(payload) {
  try {
    var sheet = getMainSheet();
    var timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    
    payload.items.forEach(function(item) {
      // H열(8): 담당자 이름, I열(9): 반영 여부, J열(10): 반영 제출일시
      sheet.getRange(item.rowIndex, 8).setValue(payload.repName);
      sheet.getRange(item.rowIndex, 9).setValue(item.reflection);
      sheet.getRange(item.rowIndex, 10).setValue(timestamp);
    });
    
    return { success: true };
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}
