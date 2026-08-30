// 1. 웹 접속 처리 (보안 라우팅 및 새 대시보드 추가)
function doGet(e) {
  var mode = e.parameter ? e.parameter.mode : null;

  if (mode === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').setTitle('관리자용 제출 내역 확인').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // ★ 추가된 학교 반영 현황 관리자 대시보드
  if (mode === 'admin_school') {
    return HtmlService.createHtmlOutputFromFile('AdminSchool').setTitle('학교 반영 현황 대시보드').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (mode === 'school') {
    return HtmlService.createHtmlOutputFromFile('School').setTitle('학교 담당자 컨설팅 반영 시스템').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (mode === 'expert') {
    return HtmlService.createHtmlOutputFromFile('Index').setTitle('중학교 교육과정 컨설팅 입력 시스템').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  var errorHtml = '<div style="text-align:center; margin-top:50px; font-family:sans-serif;"><h2 style="color:#d32f2f;">접근 권한이 없습니다 (403 Forbidden)</h2><p>올바른 접속 주소(URL)를 확인하신 후 다시 접속해주세요.</p></div>';
  return HtmlService.createHtmlOutput(errorHtml).setTitle('잘못된 접근').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function extractFileId(url) {
  if (!url) return null;
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

function getMainSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// 2. [위원용] 데이터 저장 로직
function submitData(data) {
  try {
    var sheet = getMainSheet();
    var folderId = "1EbWiQrkiHadHAm_aw0vZY5i8xmAndZrD"; 
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

// 3. [관리자용] 전체 데이터 그룹화 (H, I, J열 포함 반영)
function getAllDataForAdmin() {
  var sheet = getMainSheet();
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var ts = data[i][0];
    var consultant = data[i][1];
    var school = data[i][2];
    var key = ts + "_" + consultant + "_" + school;
    
    if (!groups[key]) {
      groups[key] = { 
        id: key, timestamp: ts, consultantName: consultant, schoolName: school, 
        repName: data[i][7] || "", schoolSubmitTime: data[i][9] || "", items: [] 
      };
    }
    // 데이터 보정 (첫 행에 담당자가 없어도 같은 그룹의 다른 행에 있으면 병합)
    if (data[i][7]) groups[key].repName = data[i][7];
    if (data[i][9]) groups[key].schoolSubmitTime = data[i][9];

    groups[key].items.push({ 
      grade: data[i][3], type: data[i][4], content: data[i][5], 
      imageUrl: data[i][6], reflection: data[i][8] || ""
    });
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

// 4. [학교 담당자용] 나이스 번호 조회 및 저장
function getSchoolDataByNeis(neisCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mappingSheet = ss.getSheetByName('학교코드');
  if (!mappingSheet) throw new Error("'학교코드' 시트가 존재하지 않습니다.");
  var mappingData = mappingSheet.getDataRange().getDisplayValues();
  var schoolName = "";
  for(var i=1; i<mappingData.length; i++) {
    if(mappingData[i][0].replace(/\s/g, '') === neisCode.replace(/\s/g, '')) {
      schoolName = mappingData[i][1]; break;
    }
  }
  if(!schoolName) throw new Error("일치하는 나이스 번호가 없습니다.");
  
  var mainSheet = getMainSheet();
  var mainData = mainSheet.getDataRange().getDisplayValues();
  var items = [];
  var repName = ""; 
  for(var i=1; i<mainData.length; i++) {
    if(mainData[i][2] === schoolName) {
      items.push({ rowIndex: i + 1, grade: mainData[i][3], type: mainData[i][4], content: mainData[i][5], imageUrl: mainData[i][6], reflection: mainData[i][8] || "" });
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
      sheet.getRange(item.rowIndex, 8).setValue(payload.repName);
      sheet.getRange(item.rowIndex, 9).setValue(item.reflection);
      sheet.getRange(item.rowIndex, 10).setValue(timestamp);
    });
    return { success: true };
  } catch(error) { return { success: false, error: error.toString() }; }
}

// [추가] 현재 웹 앱의 진짜 기본 주소(URL)를 가져오는 함수
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
