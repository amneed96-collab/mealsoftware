/**
 * মেস ব্যবস্থাপনা - Google Apps Script ব্যাকএন্ড
 * সেটআপ নির্দেশনা html ফাইলের সাথে দেওয়া আছে।
 */

var SHEET_MEMBERS = 'Members';
var SHEET_DEPOSITS = 'Deposits';
var SHEET_MEALS = 'MealEntries';
var SHEET_MENU = 'MenuItems';
var SHEET_DAILYMENU = 'DailyMenu';

var MEMBER_HEADERS = ['ID','নাম','মোবাইল','ধরন','ক্লাস','তারিখ','ঠিকানা','ছবি'];
var DEPOSIT_HEADERS = ['ID','সদস্য_আইডি','ভর্তি_ফি','ভাড়ার_মাস','ভাড়ার_বছর','ভাড়ার_পরিমাণ','তারিখ'];
var MENU_HEADERS = ['ID','নাম','ছবি'];
var DAILYMENU_HEADERS = ['তারিখ','সকাল','দুপুর','রাত'];

function getSS(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name, headers){
  var ss = getSS();
  var sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normalizeDateValue(val){
  if(val instanceof Date){
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return val;
}

function sheetToObjects(sh){
  var data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  var headers = data[0];
  var out = [];
  for(var r=1;r<data.length;r++){
    if(data[r].join('') === '') continue;
    var obj = {};
    for(var i=0;i<headers.length;i++){
      var val = data[r][i];
      if(headers[i] === 'তারিখ') val = normalizeDateValue(val);
      obj[headers[i]] = val;
    }
    out.push(obj);
  }
  return out;
}

function getMealsStructured(sh){
  var data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  var headers = data[0];
  var out = [];
  for(var r=1;r<data.length;r++){
    if(data[r][0] === '' || data[r][0] === null) continue;
    var entry = { date: normalizeDateValue(data[r][0]), meals: {} };
    for(var i=1;i<headers.length;i++){
      var h = String(headers[i]);
      var parts = h.split(' - ');
      if(parts.length < 2) continue;
      var type = parts[parts.length-1];
      var name = parts.slice(0, parts.length-1).join(' - ');
      if(!entry.meals[name]) entry.meals[name] = { 'সকাল':0, 'দুপুর':0, 'রাত':0, 'মোট':0 };
      entry.meals[name][type] = data[r][i] || 0;
    }
    out.push(entry);
  }
  return out;
}

function doGet(e){
  var action = e.parameter.action;
  var membersSh = getSheet(SHEET_MEMBERS, MEMBER_HEADERS);
  var depositsSh = getSheet(SHEET_DEPOSITS, DEPOSIT_HEADERS);
  var mealsSh = getSheet(SHEET_MEALS, ['তারিখ']);
  var menuSh = getSheet(SHEET_MENU, MENU_HEADERS);
  var dailyMenuSh = getSheet(SHEET_DAILYMENU, DAILYMENU_HEADERS);

  if(action === 'getAll'){
    return jsonOut({
      ok:true,
      members: sheetToObjects(membersSh),
      deposits: sheetToObjects(depositsSh),
      meals: getMealsStructured(mealsSh),
      menuItems: sheetToObjects(menuSh),
      dailyMenus: sheetToObjects(dailyMenuSh)
    });
  }
  return jsonOut({ok:false, error:'unknown action'});
}

function doPost(e){
  var body;
  try{ body = JSON.parse(e.postData.contents); }
  catch(err){ return jsonOut({ok:false, error:'invalid json'}); }

  var action = body.action;
  var membersSh = getSheet(SHEET_MEMBERS, MEMBER_HEADERS);
  var depositsSh = getSheet(SHEET_DEPOSITS, DEPOSIT_HEADERS);
  var mealsSh = getSheet(SHEET_MEALS, ['তারিখ']);
  var menuSh = getSheet(SHEET_MENU, MENU_HEADERS);
  var dailyMenuSh = getSheet(SHEET_DAILYMENU, DAILYMENU_HEADERS);

  try{
    if(action === 'saveMember') return jsonOut(saveRow(membersSh, body.data, MEMBER_HEADERS));
    if(action === 'deleteMember') return jsonOut(deleteRowById(membersSh, body.id));
    if(action === 'saveDeposit') return jsonOut(saveRow(depositsSh, body.data, DEPOSIT_HEADERS));
    if(action === 'deleteDeposit') return jsonOut(deleteRowById(depositsSh, body.id));
    if(action === 'saveMealEntry') return jsonOut(saveMealEntry(mealsSh, body.date, body.meals));
    if(action === 'deleteMealEntry') return jsonOut(deleteRowByDate(mealsSh, body.date));
    if(action === 'saveMenuItem') return jsonOut(saveRow(menuSh, body.data, MENU_HEADERS));
    if(action === 'deleteMenuItem') return jsonOut(deleteRowById(menuSh, body.id));
    if(action === 'saveDailyMenu') return jsonOut(saveDailyMenuRow(dailyMenuSh, body.date, body.values));
    if(action === 'deleteDailyMenu') return jsonOut(deleteRowByDate(dailyMenuSh, body.date));
    return jsonOut({ok:false, error:'unknown action'});
  } catch(err){
    return jsonOut({ok:false, error: err.message});
  }
}

function saveRow(sh, data, headers){
  var values = sh.getDataRange().getValues();
  for(var r=1;r<values.length;r++){
    if(String(values[r][0]) === String(data.ID)){
      var rowNum = r+1;
      for(var i=0;i<headers.length;i++){
        sh.getRange(rowNum, i+1).setValue(data[headers[i]] !== undefined ? data[headers[i]] : '');
      }
      return {ok:true, updated:true};
    }
  }
  var newRow = [];
  for(var j=0;j<headers.length;j++){ newRow.push(data[headers[j]] !== undefined ? data[headers[j]] : ''); }
  sh.appendRow(newRow);
  return {ok:true, updated:false};
}

function deleteRowById(sh, id){
  var values = sh.getDataRange().getValues();
  for(var r=1;r<values.length;r++){
    if(String(values[r][0]) === String(id)){
      sh.deleteRow(r+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'not found'};
}

function deleteRowByDate(sh, date){
  var data = sh.getDataRange().getValues();
  for(var r=1;r<data.length;r++){
    var cellStr = normalizeDateValue(data[r][0]);
    if(String(cellStr) === String(date)){
      sh.deleteRow(r+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'not found'};
}

function saveMealEntry(sh, date, meals){
  // meals shape: { "সদস্যর নাম": { সকাল, দুপুর, রাত, মোট } }
  var types = ['সকাল','দুপুর','রাত','মোট'];
  var data = sh.getDataRange().getValues();
  var headers = data[0] || ['তারিখ'];
  if(headers.length === 0){ sh.appendRow(['তারিখ']); headers = ['তারিখ']; }

  var namesToAdd = [];
  for(var name in meals){
    types.forEach(function(type){
      var key = name + ' - ' + type;
      if(headers.indexOf(key) === -1) namesToAdd.push(key);
    });
  }
  namesToAdd.forEach(function(key){
    headers.push(key);
    sh.getRange(1, headers.length).setValue(key);
  });

  data = sh.getDataRange().getValues();
  headers = data[0];

  var rowIndex = -1;
  for(var r=1;r<data.length;r++){
    var cellStr = normalizeDateValue(data[r][0]);
    if(String(cellStr) === String(date)){ rowIndex = r+1; break; }
  }
  if(rowIndex === -1){
    sh.appendRow([date]);
    rowIndex = sh.getLastRow();
  }

  for(var name2 in meals){
    types.forEach(function(type){
      var key = name2 + ' - ' + type;
      var col = headers.indexOf(key) + 1;
      sh.getRange(rowIndex, col).setValue(meals[name2][type] || 0);
    });
  }
  return {ok:true};
}

function saveDailyMenuRow(sh, date, values){
  // values shape: { সকাল, দুপুর, রাত } — প্রতিটি একটি মেনু আইটেমের ID অথবা ফাঁকা স্ট্রিং
  var data = sh.getDataRange().getValues();
  var rowIndex = -1;
  for(var r=1;r<data.length;r++){
    var cellStr = normalizeDateValue(data[r][0]);
    if(String(cellStr) === String(date)){ rowIndex = r+1; break; }
  }
  var rowValues = [date, values['সকাল']||'', values['দুপুর']||'', values['রাত']||''];
  if(rowIndex === -1){
    sh.appendRow(rowValues);
  } else {
    sh.getRange(rowIndex, 1, 1, 4).setValues([rowValues]);
  }
  return {ok:true};
}
