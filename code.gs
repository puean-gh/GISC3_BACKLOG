function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('รวม GoogleSheet ต่างๆ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData(fileId, sheetIndex) {
  var ss = SpreadsheetApp.openById(fileId);
  var sheets = ss.getSheets();
  var sheet = sheets[sheetIndex || 0]; 
  var data = sheet.getDataRange().getDisplayValues(); 
  
  if (fileId === "1PKMVSFCgBwU42JRLWflpJa89zG2EtZ0l43FSuDppfhY") {
    data = data.map(function(row) { return row.slice(0, 11); });
  } 
  else if (fileId === "1zYbB18hl0hJ7kNY4lBSXVF53m0ZtXLE7cYCV4tKc2lI") {
    data = data.map(function(row) { return row.slice(0, 15); });
  } 
  else if (fileId === "1MeYS4aTMpX5S7DE56whS21rPIp7Wa7-KsFm3nu0yd0g") {
    data = data.map(function(row) { return row.slice(0, 6); });
  } 
  else if (fileId === "1JgzqeXHKB_n5gWZu9hZgcU6RbgckSls3xdM7Nvf6U90") {
    data = data.map(function(row) { return row.slice(0, 4); });
  } 
  else if (fileId === "1PwQbHXBAQlpdWlRS0TiNoIMeDblaDBZd_ZhnqiOxDYk") {
    data = data.map(function(row) { return row.slice(0, 6); });
  }

  var sheetNames = sheets.map(function(s) { return s.getName(); });
  
  return {
    values: data,
    sheetName: ss.getName(),
    currentSheetName: sheet.getName(),
    allSheets: sheetNames,
    fileId: fileId
  };
}

function getDashboardSummary() {
  var files = getFileList();
  var summary = [];
  
  for (var i = 0; i < files.length; i++) {
    var fileId = files[i].id;
    var name = files[i].name;
    var totalRows = 0;
    
    try {
      var ss = SpreadsheetApp.openById(fileId);
      var sheet = ss.getSheets()[0];
      
      if (fileId === "1zYbB18hl0hJ7kNY4lBSXVF53m0ZtXLE7cYCV4tKc2lI") {
        var s1 = ss.getSheets()[0].getLastRow() - 1; 
        var s2 = ss.getSheets()[1].getLastRow() - 1;
        totalRows = (s1 > 0 ? s1 : 0) + (s2 > 0 ? s2 : 0);
      } 
      // เงื่อนไขใหม่: กรองค่าว่างของ OPSAonGIS ไม่ให้นับเข้าการ์ด
      else if (fileId === "1PwQbHXBAQlpdWlRS0TiNoIMeDblaDBZd_ZhnqiOxDYk") {
        var data = sheet.getDataRange().getDisplayValues();
        if (data.length > 1) {
            var headers = data[0];
            var targetCol = -1;
            for (var j = 0; j < headers.length; j++) {
                if (headers[j].toString().trim() === "กฟฟ.") { targetCol = j; break; }
            }
            
            for (var r = 1; r < data.length; r++) {
                var val = targetCol !== -1 ? data[r][targetCol].toString().trim() : data[r][0].toString().trim();
                if (val !== "" && val !== "-") {
                    totalRows++;
                }
            }
        }
      } 
      else {
        var s = sheet.getLastRow() - 1;
        totalRows = (s > 0 ? s : 0);
      }
    } catch (e) {
      totalRows = "Error";
    }
    
    summary.push({ name: name, count: totalRows });
  }
  
  return summary;
}

function getFileList() {
  return [
    { name: "BackLog_ADS", id: "1PKMVSFCgBwU42JRLWflpJa89zG2EtZ0l43FSuDppfhY" },
    { name: "BackLog_ISU", id: "1zYbB18hl0hJ7kNY4lBSXVF53m0ZtXLE7cYCV4tKc2lI" },
    { name: "Waive มิเตอร์", id: "1MeYS4aTMpX5S7DE56whS21rPIp7Wa7-KsFm3nu0yd0g" },
    { name: "Waive หม้อแปลง", id: "1JgzqeXHKB_n5gWZu9hZgcU6RbgckSls3xdM7Nvf6U90" },
    { name: "ค่าผิดปกติ จาก OPSAonGIS", id: "1PwQbHXBAQlpdWlRS0TiNoIMeDblaDBZd_ZhnqiOxDYk" }
  ];
}
