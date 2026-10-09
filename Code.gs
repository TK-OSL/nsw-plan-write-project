function doGet() {
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail;
  template.userRole = userRole;

  return template.evaluate()
      .setTitle('ระบบเขียนโครงการโรงเรียน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserRole(email) {
  if (!email) return 'none';
  try {
    var ss = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II');
    var userSheet = ss.getSheetByName('Users');
    if (!userSheet) return 'none';
    var data = userSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        return data[i][1].toString().trim().toLowerCase();
      }
    }
  } catch(e) {}
  return 'none';
}

function syncData() {
  try {
    var sourceId = '14z4vhoBP-uTWs1gObzqNdNwSq1QTiAeQpfwPOlq2cCE';
    var destId = '1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II';
    var sourceSS = SpreadsheetApp.openById(sourceId);
    var destSS = SpreadsheetApp.openById(destId);

    var sourceSheet1 = sourceSS.getSheetByName('Stat');
    var destSheet1 = destSS.getSheetByName('Project');
    if (sourceSheet1 && destSheet1) {
      var destMaxRows1 = destSheet1.getMaxRows();
      if (destMaxRows1 > 1) destSheet1.getRange(2, 1, destMaxRows1 - 1, 9).clearContent(); 
      var lastRow1 = sourceSheet1.getLastRow();
      if (lastRow1 > 1) {
        var rawData1 = sourceSheet1.getRange(2, 1, lastRow1 - 1, 8).getValues();
        var validData1 = rawData1.filter(function(row) { return row[0] !== null && row[0].toString().trim() !== ""; });
        if (validData1.length > 0) {
          var d = new Date(); var year = d.getFullYear() + 543; var month = ("0" + (d.getMonth() + 1)).slice(-2);
          var planIds = [];
          for (var i = 0; i < validData1.length; i++) planIds.push(["Plan" + year + month + ("000" + (i + 1)).slice(-3)]);
          destSheet1.getRange(2, 1, planIds.length, 1).setValues(planIds);
          destSheet1.getRange(2, 2, validData1.length, validData1[0].length).setValues(validData1);
        }
      }
    } else throw new Error('ไม่พบชีต Stat หรือ Project');

    var sourceSheet2 = sourceSS.getSheetByName('Data');
    var destSheet2 = destSS.getSheetByName('กิจกรรม');
    if (sourceSheet2 && destSheet2) {
      var destMaxRows2 = destSheet2.getMaxRows();
      if (destMaxRows2 > 1) destSheet2.getRange(2, 1, destMaxRows2 - 1, 14).clearContent();
      var lastRow2 = sourceSheet2.getLastRow();
      if (lastRow2 > 1) {
        var rawData2 = sourceSheet2.getRange(2, 1, lastRow2 - 1, 14).getValues();
        var validData2 = rawData2.filter(function(row) { return row.some(function(cell) { return cell !== null && cell.toString().trim() !== ""; }); });
        if (validData2.length > 0) destSheet2.getRange(2, 1, validData2.length, validData2[0].length).setValues(validData2);
      }
    } else throw new Error('ไม่พบชีต Data หรือ กิจกรรม');
    return { success: true, message: 'ดึงข้อมูลสำเร็จ!' };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function getDropdownData() {
  var ss = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II');
  function getColA(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if(!sheet) return [];
    var lastRow = sheet.getLastRow();
    if(lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().filter(function(v) { return v.trim() !== ""; });
  }
  
  function getUsers() {
    var sheet = ss.getSheetByName('Users');
    if(!sheet) return [];
    var lastRow = sheet.getLastRow();
    if(lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    var users = [];
    for(var i=0; i<data.length; i++) {
      if(data[i][0].trim() !== "") users.push({ email: data[i][0], role: data[i][1] });
    }
    return users;
  }

  var teacherSheet = ss.getSheetByName('ฐานข้อมูลครู');
  var teacherNames = [], teacherPositions = [], directorName = "", directorPos = "";
  if(teacherSheet && teacherSheet.getLastRow() > 1) {
    var tData = teacherSheet.getRange(2, 2, teacherSheet.getLastRow() - 1, 6).getDisplayValues(); 
    var posSet = new Set();
    for(var i=0; i<tData.length; i++) {
      var name = tData[i][0].trim(), pos = tData[i][5].trim();
      if(name) teacherNames.push(name);
      if(pos) posSet.add(pos);
      if(pos.indexOf("ผู้อำนวยการโรงเรียน") > -1 || (pos.indexOf("ผู้อำนวยการ") > -1 && pos.indexOf("รอง") === -1)) {
        directorName = name; directorPos = pos;
      }
    }
    teacherPositions = Array.from(posSet);
  }

  var groupLeaderPos = getColA('ตำแหน่งหัวหน้ากลุ่ม');
  var combinedPositions = Array.from(new Set(teacherPositions.concat(groupLeaderPos)));

  return {
    strategies: getColA('ฐานข้อมูลกลยุทธ์โรงเรียน'), stdObec: getColA('ฐานข้อมูลมาตรฐาน สพฐ'), stdDream: getColA('ฐานข้อมูลมาตรฐาน รร ในฝัน'),
    groupLeaderPos: groupLeaderPos, users: getUsers(),
    teacherNames: teacherNames, teacherPositions: combinedPositions, directorName: directorName, directorPos: directorPos
  };
}

function getProjectList() {
  var data = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName('Project').getDataRange().getDisplayValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].trim() !== "" && data[i][1] && data[i][1].trim() !== "") projects.push({ planId: data[i][0].trim(), name: data[i][1].trim() });
  }
  return projects;
}

function getProjectDetails(planId) {
  var ss = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II');
  var projData = ss.getSheetByName('Project').getDataRange().getValues();
  var projRow = null;
  for (var i = 1; i < projData.length; i++) { if (projData[i][0] === planId) { projRow = projData[i]; projRow.rowIndex = i + 1; break; } }
  if (!projRow) return null;

  var actData = ss.getSheetByName('กิจกรรม').getDataRange().getValues();
  var fYear = projRow[9]||"", fDept = projRow[12]||"", fSubj = projRow[13]||"", fMgr = projRow[14]||"";
  
  var defaultActivities = [];
  for (var a = 1; a < actData.length; a++) {
    if (actData[a][4] === projRow[1]) {
      if(!fYear && actData[a][1]) fYear = actData[a][1]; if(!fDept && actData[a][2]) fDept = actData[a][2];
      if(!fSubj && actData[a][3]) fSubj = actData[a][3]; if(!fMgr && actData[a][5]) fMgr = actData[a][5]; 
      var actName = actData[a][7] || ""; 
      var actBudget = actData[a][9] || actData[a][8] || 0; 
      var actRes = actData[a][6] || ""; 
      if(actName) defaultActivities.push([actName, "", "", "", actBudget, actRes]);
    }
  }
  
  var startDate = (projRow[10] instanceof Date) ? Utilities.formatDate(projRow[10], "GMT+7", "yyyy-MM-dd") : "";
  var endDate = (projRow[11] instanceof Date) ? Utilities.formatDate(projRow[11], "GMT+7", "yyyy-MM-dd") : "";
  if(fYear && !isNaN(parseInt(fYear))) {
    var yearCE = parseInt(fYear) - 543;
    if(!startDate) startDate = (yearCE - 1) + "-10-01";
    if(!endDate) endDate = yearCE + "-09-30";
  }

  var subSheetsNames = ['กลยุทธ์', 'มาตรฐาน รร ในฝัน', 'มาตรฐาน สพฐ', 'หลักการ', 'วัตถุประสงค์', 'เป้าหมายเชิงปริมาณ', 'เป้าหมายเชิงคุณภาพ', 'การดำเนินงาน', 'หน่วยงานที่เกี่ยวข้อง', 'ผลที่คาดว่าจะได้รับ', 'การนิเทศติดตามและประเมินผล'];
  var subData = {};
  subSheetsNames.forEach(function(sName) {
    subData[sName] = [];
    var sheet = ss.getSheetByName(sName);
    if(sheet) {
      var vals = sheet.getDataRange().getDisplayValues();
      for(var r = 1; r < vals.length; r++) {
        if(vals[r][0] === planId) subData[sName].push(vals[r].slice(1));
      }
    }
  });

  if (subData['การดำเนินงาน'].length === 0 && defaultActivities.length > 0) subData['การดำเนินงาน'] = defaultActivities;

  return {
    row: projRow.rowIndex, planId: projRow[0], name: projRow[1], 
    year: fYear, department: fDept, subjectArea: fSubj, manager: fMgr, startDate: startDate, endDate: endDate,
    charact: projRow[15]||"", propName: projRow[16]||"", propPos: projRow[17]||"", groupLeaderName: projRow[18]||"", groupLeaderPos: projRow[19]||"", 
    app1Name: projRow[20]||"", app1Pos: projRow[21]||"", app2Name: projRow[22]||"", app2Pos: projRow[23]||"", app3Name: projRow[24]||"", app3Pos: projRow[25]||"", finalName: projRow[26]||"", finalPos: projRow[27]||"", status: projRow[28] || "รออนุมัติ",
    budget: { approved: projRow[3]||0, udnun: projRow[4]||0, patana: projRow[5]||0, raidai: projRow[6]||0, other: projRow[7]||0, total: projRow[8]||0 },
    subData: subData
  };
}

function saveProjectData(form) {
  try {
    var ss = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II');
    var planId = form.planId;
    var sheetProj = ss.getSheetByName('Project');
    var sDate = form.startDate ? new Date(form.startDate) : "";
    var eDate = form.endDate ? new Date(form.endDate) : "";
    
    // บันทึกเฉพาะข้อมูลหลักที่อยู่ใน Project ตั้งแต่ J ถึง AC
    var updateValues = [[
      form.year, sDate, eDate, form.department, form.subjectArea, form.manager, form.charact,
      form.propName, form.propPos, form.groupLeaderName, form.groupLeaderPos,
      form.app1Name, form.app1Pos, form.app2Name, form.app2Pos, 
      form.app3Name, form.app3Pos, form.finalName, form.finalPos, form.status
    ]];
    sheetProj.getRange(parseInt(form.row), 10, 1, 20).setValues(updateValues);

    var subSheets = ['กลยุทธ์', 'มาตรฐาน รร ในฝัน', 'มาตรฐาน สพฐ', 'หลักการ', 'วัตถุประสงค์', 'เป้าหมายเชิงปริมาณ', 'เป้าหมายเชิงคุณภาพ', 'การดำเนินงาน', 'หน่วยงานที่เกี่ยวข้อง', 'ผลที่คาดว่าจะได้รับ', 'การนิเทศติดตามและประเมินผล'];
    
    subSheets.forEach(function(sName) {
      var sheet = ss.getSheetByName(sName);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        var maxCols = sheet.getMaxColumns();
        var vals = lastRow > 0 ? sheet.getRange(1, 1, lastRow, maxCols).getValues() : [];
        var newData = [];
        
        // 1. เก็บหัวตาราง และ โครงการอื่นที่ไม่ได้กำลังแก้ (กันบั๊กชีตว่างเปล่า)
        for (var i = 0; i < vals.length; i++) {
          if (i === 0 || vals[i][0] !== planId) { 
            if (i === 0 || vals[i].join("").trim() !== "") newData.push(vals[i]);
          }
        }

        // ถ้าชีตพัง ไม่มี Header ให้สร้างจำลองขึ้นมา 1 แถวเพื่อกัน Error
        if (newData.length === 0) {
          var dummyHeader = ["PlanID"];
          for(var c=1; c<maxCols; c++) dummyHeader.push("Col"+c);
          newData.push(dummyHeader);
        }
        
        // 2. นำข้อมูลที่ User กรอก มาแนบ PlanID เข้าไป (ถ้ามีค่าว่าง ไม่ต้องใส่)
        if (form.subData[sName] && form.subData[sName].length > 0) {
          form.subData[sName].forEach(function(itemArray) {
            var isEmpty = true;
            for(var j=0; j<itemArray.length; j++) {
              if(itemArray[j] !== null && itemArray[j] !== undefined && itemArray[j].toString().trim() !== "") {
                isEmpty = false; break;
              }
            }
            if (!isEmpty) {
              var rowData = [planId].concat(itemArray);
              var headerLen = newData[0].length;
              while(rowData.length < headerLen) rowData.push(""); 
              if(rowData.length > headerLen) rowData = rowData.slice(0, headerLen);
              newData.push(rowData);
            }
          });
        }
        
        // 3. ล้างชีตแล้วเขียนทับ
        sheet.clearContents();
        if(newData.length > 0) {
          sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
        }
      }
    });

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ----------------------------------------------------
// ระบบแก้ไข/เรียงลำดับ ข้อมูลฐานข้อมูลทั้งหมด
// ----------------------------------------------------
function updateMasterDataList(sheetName, newList) {
  try {
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName(sheetName);
    var lastRow = sheet.getLastRow();
    if(lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
    if(newList && newList.length > 0) {
      var arr = newList.map(function(item) { return [item]; });
      sheet.getRange(2, 1, arr.length, 1).setValues(arr);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function addMasterData(sheetName, value) {
  try {
    SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName(sheetName).appendRow([value.trim()]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteMasterData(sheetName, value) {
  try {
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName(sheetName);
    var data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].trim() === value.trim()) { sheet.deleteRow(i + 2); return { success: true }; }
    }
  } catch(e) { return { success: false, message: e.toString() }; }
}

function editMasterData(sheetName, oldValue, newValue) {
  try {
    if (!newValue || newValue.trim() === "") throw new Error("ข้อมูลว่างเปล่า");
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName(sheetName);
    var data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].trim() === oldValue.trim()) { 
        sheet.getRange(i + 2, 1).setValue(newValue.trim()); return { success: true }; 
      }
    }
    throw new Error("ไม่พบข้อมูลเดิม");
  } catch(e) { return { success: false, message: e.toString() }; }
}

function addUser(email, role) {
  try {
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName('Users');
    sheet.appendRow([email.trim(), role.trim()]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function editUser(oldEmail, newEmail, newRole) {
  try {
    if (!newEmail || newEmail.trim() === "") throw new Error("อีเมลว่างเปล่า");
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName('Users');
    var data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 2).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].trim() === oldEmail.trim()) { 
        sheet.getRange(i + 2, 1, 1, 2).setValues([[newEmail.trim(), newRole.trim()]]); 
        return { success: true }; 
      }
    }
    throw new Error("ไม่พบข้อมูลเดิม");
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteUser(email) {
  try {
    var sheet = SpreadsheetApp.openById('1aYlTkFSZs9G3qjFonPP_5gv6E547S8XHfy04OrBY-II').getSheetByName('Users');
    var data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].trim() === email.trim()) { sheet.deleteRow(i + 2); return { success: true }; }
    }
  } catch(e) { return { success: false, message: e.toString() }; }
}
