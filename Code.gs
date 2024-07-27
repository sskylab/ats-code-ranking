// FIXED: Add Google Sheet by name sheet over here!!
var folderID = "1-_cM-Dj0TMHQL2Rbnj491rA1p-og6e5y";
var baseSheetId = SpreadsheetApp.openById('14K5I2fxSC_nPMvooJiXx7H0xE0OckqJnNuQ_N3_E2FA')
var sheet_user_permission = baseSheetId.getSheetByName('user')
var sheet_quiz_dotNetCSharp = baseSheetId.getSheetByName('Quiz-C#')
var sheet_quiz_XML = baseSheetId.getSheetByName('Quiz-XML')
var sheet_quiz_SQL = baseSheetId.getSheetByName('Quiz-SQL')
var sheet_quiz_Typescript = baseSheetId.getSheetByName('Quiz-Typescript')
var sheet_quiz_React = baseSheetId.getSheetByName('Quiz-React')
var sheet_quiz_Kotlin = baseSheetId.getSheetByName('Quiz-Kotlin')
var sheet_quiz_PostgreSQL = baseSheetId.getSheetByName('Quiz-PostgreSQL')
var sheet_quiz_Python = baseSheetId.getSheetByName('Quiz-Python')
var sheet_quiz_bootstrap5 = baseSheetId.getSheetByName('Quiz-bootstrap5')
var sheet_quiz_PHP = baseSheetId.getSheetByName('Quiz-PHP')
var sheet_quiz_jQuery = baseSheetId.getSheetByName('Quiz-jQuery')
var sheet_quiz_CPlus = baseSheetId.getSheetByName('Quiz-CPlus')
var sheet_quiz_CSS = baseSheetId.getSheetByName('Quiz-CSS')
var sheet_quiz_HTML = baseSheetId.getSheetByName('Quiz-HTML')
var sheet_quiz_Javascript = baseSheetId.getSheetByName('Quiz-Javascript')
var sheet_quiz_JAVA = baseSheetId.getSheetByName('Quiz-JAVA')
var sheet_ATS_Support = baseSheetId.getSheetByName('support')

function doGet(e) {
  let page = e.parameter.page;
  if (page == null) page = "login";

  var output = HtmlService.createTemplateFromFile(getCurrentPage(page));
  return output.evaluate()
    .setTitle('ATS Code Ranking')
    .setFaviconUrl('https://awareth.aware-cdn.net/wp-content/uploads/A-Logo-512X512px.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

// FIXED: SET/GET APPLICATION VERSION WHEN WE WILL DEPLOYMENT
function getAppVersion() {
  var appVersion = '1.1.5';
  return appVersion;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

function getUrl() {
  var url = ScriptApp.getService().getUrl()
  return url
}

function getCurrentPage(page) {
  var cache = CacheService.getUserCache();
  var sessionToken = cache.get('sessionToken');

  if (sessionToken) {
    if (page == 'login') {
      return 'main';
    }
    else {
      return page
    }
  }
  else {
    return "login";
  }
}

function login(email) {
  var columnRange = sheet_user_permission.getRange("C:C");
  var idList = columnRange.getValues();
  var valuesArray = [];

  for (var i = 0; i < idList.length; i++) {
    if (email.toLowerCase() == idList[i][0].toLowerCase()) {
      var rowValues = sheet_user_permission.getRange(i + 1, 1, 1, sheet_user_permission.getLastColumn()).getValues()[0];
      var valuesObj = {
        userID: rowValues[0],
        userName: rowValues[1],
        userEmail: rowValues[2],
        position: rowValues[3],
        levelCode: rowValues[4],
        userImageProfile: getImageUrlByName(rowValues[0])
      };

      valuesArray.push(valuesObj);

      // Create a session token and store it in the user's cache
      var sessionToken = Utilities.getUuid();
      var cache = CacheService.getUserCache();
      cache.put('sessionToken', sessionToken, 3600); // Session expires in 1 hour

      return { success: true, message: 'credentials success', sessionToken: sessionToken, data: valuesArray };
    }
  }

  return { success: false, message: 'Invalid credentials', data: null };
}

function checkSession() {
  var cache = CacheService.getUserCache();
  var sessionToken = cache.get('sessionToken');

  if (sessionToken) {
    return { loggedIn: true };
  } else {
    return { loggedIn: false };
  }
}

function logout() {
  var cache = CacheService.getUserCache();
  cache.remove('sessionToken');
  return { success: true };
}

function getImageUrlByName(fileName) {
  var folder = DriveApp.getFolderById(folderID);
  var files = folder.getFilesByName(fileName);

  if (!files.hasNext()) {
    return fileUrl = 'https://conndv.aware.co.th:8080/ats-happy/v1/download/file?path=/files/ats-share/ATS_AVATAR/&file=avatar_default.png';
  }

  var file = files.next();

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileUrl = 'https://drive.google.com/thumbnail?id=' + file.getId();

  return fileUrl;
}

function uploadFileToDrive(base64Data, fileName) {
  var contentType = base64Data.substring(0, base64Data.indexOf(','));
  var byteString = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(byteString, contentType, fileName);

  var folder = DriveApp.getFolderById(folderID);
  var existingFiles = folder.getFilesByName(fileName);

  if (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function appendATSCodeRankingSupportData(values) {
  var columnRange = sheet_ATS_Support.getRange("B:B");

  sheet_ATS_Support.appendRow(
    [
      new Date().getTime().toString(), //support_id
      new Date(), //date
      values.fullName, //full_name
      values.email, //email
      values.telephone, // telephone
      values.textArea, // desc

    ]
  );

  return { success: true, message: 'success' };
}

function convertStringToDate(dateString) {
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-based
  const day = String(date.getDate()).padStart(2, '0');

  const formattedDate = `${year}-${month}-${day}`;

  return formattedDate.toString()
}

function getUserProfile(empCode) {
  var columnRange = sheet_user_permission.getRange("A:A");
  var idList = columnRange.getValues();

  for (var i = 0; i < idList.length; i++) {
    if (empCode.toString() == idList[i][0].toString()) {
      var userValues = sheet_user_permission.getRange(i + 1, 1, 1, sheet_user_permission.getLastColumn()).getValues()[0];
      var valuesObj = {
        userID: userValues[0],
        userName: userValues[1],
        userEmail: userValues[2],
        imageProfile: userValues[5]
      };

      return valuesObj;
    }
  }
}

function getYourScoreBoard(empCodeStr) {
  var dotNet_List = sheet_quiz_dotNetCSharp.getRange("AB:AB").getValues();
  var xml_List = sheet_quiz_XML.getRange("AB:AB").getValues();
  var sql_List = sheet_quiz_SQL.getRange("AB:AB").getValues();
  var typescript_List = sheet_quiz_Typescript.getRange("AB:AB").getValues();
  var react_List = sheet_quiz_React.getRange("AB:AB").getValues();
  var kotlin_List = sheet_quiz_Kotlin.getRange("AB:AB").getValues();
  var postgreSql_List = sheet_quiz_PostgreSQL.getRange("AB:AB").getValues();
  var python_List = sheet_quiz_Python.getRange("AB:AB").getValues();
  var bootstrap5_List = sheet_quiz_bootstrap5.getRange("AB:AB").getValues();
  var php_List = sheet_quiz_PHP.getRange("AB:AB").getValues();
  var jQuery_List = sheet_quiz_jQuery.getRange("AB:AB").getValues();
  var cPlus_List = sheet_quiz_CPlus.getRange("AB:AB").getValues();
  var css_List = sheet_quiz_CSS.getRange("AB:AB").getValues();
  var javascript_List = sheet_quiz_Javascript.getRange("AB:AB").getValues();
  var java_List = sheet_quiz_JAVA.getRange("AB:AB").getValues();
  var html_List = sheet_quiz_HTML.getRange("AQ:AQ").getValues(); //42,43

  var defaultValue = [{ empEmail: "", score: "0", empCode: "", timestamp: "" }];

  let quiz_dotNet_Array = defaultValue;
  let quiz_xml_Array = defaultValue;
  let quiz_sql_Array = defaultValue;
  let quiz_typescript_Array = defaultValue;
  let quiz_react_Array = defaultValue;
  let quiz_kotlin_Array = defaultValue;
  let quiz_postgressql_Array = defaultValue;
  let quiz_python_Array = defaultValue;
  let quiz_bootstrap5_Array = defaultValue;
  let quiz_php_Array = defaultValue;
  let quiz_jQuery_Array = defaultValue;
  let quiz_cPlus_Array = defaultValue;
  let quiz_css_Array = defaultValue;
  let quiz_javascript_Array = defaultValue;
  let quiz_java_Array = defaultValue;
  let quiz_html_Array = defaultValue;

  for (var i = 0; i < dotNet_List.length; i++) {
    if (empCodeStr == dotNet_List[i][0].toString()) {
      var dotNetValues = sheet_quiz_dotNetCSharp.getRange(i + 1, 1, 1, sheet_quiz_dotNetCSharp.getLastColumn()).getValues()[0];
      var dotNetObj = {
        timestamp: dotNetValues[0].toString(),
        score: dotNetValues[1].toString(),
        empCode: dotNetValues[27].toString(),
        empEmail: dotNetValues[28].toString(),
      }
      quiz_dotNet_Array = [];
      quiz_dotNet_Array.push(dotNetObj);
    }
  }

  for (var i = 0; i < xml_List.length; i++) {
    if (empCodeStr == xml_List[i][0].toString()) {
      var xmlValues = sheet_quiz_XML.getRange(i + 1, 1, 1, sheet_quiz_XML.getLastColumn()).getValues()[0];
      var xmlObj = {
        timestamp: xmlValues[0].toString(),
        score: xmlValues[1].toString(),
        empCode: xmlValues[27].toString(),
        empEmail: xmlValues[28].toString(),
      }
      quiz_xml_Array = [];
      quiz_xml_Array.push(xmlObj);
    }
  }

  for (var i = 0; i < sql_List.length; i++) {
    if (empCodeStr == sql_List[i][0].toString()) {
      var sqlValues = sheet_quiz_SQL.getRange(i + 1, 1, 1, sheet_quiz_SQL.getLastColumn()).getValues()[0];
      var sqlObj = {
        timestamp: sqlValues[0].toString(),
        score: sqlValues[1].toString(),
        empCode: sqlValues[27].toString(),
        empEmail: sqlValues[28].toString(),
      }
      quiz_sql_Array = [];
      quiz_sql_Array.push(sqlObj);
    }
  }

  for (var i = 0; i < typescript_List.length; i++) {
    if (empCodeStr == typescript_List[i][0].toString()) {
      var typescriptValues = sheet_quiz_Typescript.getRange(i + 1, 1, 1, sheet_quiz_Typescript.getLastColumn()).getValues()[0];
      var typescriptObj = {
        timestamp: typescriptValues[0].toString(),
        score: typescriptValues[1].toString(),
        empCode: typescriptValues[27].toString(),
        empEmail: typescriptValues[28].toString(),
      }
      quiz_typescript_Array = [];
      quiz_typescript_Array.push(typescriptObj);
    }
  }

  for (var i = 0; i < react_List.length; i++) {
    if (empCodeStr == react_List[i][0].toString()) {
      var reactValues = sheet_quiz_React.getRange(i + 1, 1, 1, sheet_quiz_React.getLastColumn()).getValues()[0];
      var reactObj = {
        timestamp: reactValues[0].toString(),
        score: reactValues[1].toString(),
        empCode: reactValues[27].toString(),
        empEmail: reactValues[28].toString(),
      }
      quiz_react_Array = [];
      quiz_react_Array.push(reactObj);
    }
  }

  for (var i = 0; i < kotlin_List.length; i++) {
    if (empCodeStr == kotlin_List[i][0].toString()) {
      var kotlinValues = sheet_quiz_Kotlin.getRange(i + 1, 1, 1, sheet_quiz_Kotlin.getLastColumn()).getValues()[0];
      var kotlinObj = {
        timestamp: kotlinValues[0].toString(),
        score: kotlinValues[1].toString(),
        empCode: kotlinValues[27].toString(),
        empEmail: kotlinValues[28].toString(),
      }
      quiz_kotlin_Array = [];
      quiz_kotlin_Array.push(kotlinObj);
    }
  }

  for (var i = 0; i < postgreSql_List.length; i++) {
    if (empCodeStr == postgreSql_List[i][0].toString()) {
      var postgreSQLValues = sheet_quiz_PostgreSQL.getRange(i + 1, 1, 1, sheet_quiz_PostgreSQL.getLastColumn()).getValues()[0];
      var postgreSQLObj = {
        timestamp: postgreSQLValues[0].toString(),
        score: postgreSQLValues[1].toString(),
        empCode: postgreSQLValues[27].toString(),
        empEmail: postgreSQLValues[28].toString(),
      }
      quiz_postgressql_Array = [];
      quiz_postgressql_Array.push(postgreSQLObj);
    }
  }

  for (var i = 0; i < python_List.length; i++) {
    if (empCodeStr == python_List[i][0].toString()) {
      var pythonValues = sheet_quiz_Python.getRange(i + 1, 1, 1, sheet_quiz_Python.getLastColumn()).getValues()[0];
      var pythonObj = {
        timestamp: pythonValues[0].toString(),
        score: pythonValues[1].toString(),
        empCode: pythonValues[27].toString(),
        empEmail: pythonValues[28].toString(),
      }
      quiz_python_Array = [];
      quiz_python_Array.push(pythonObj);
    }
  }

  for (var i = 0; i < bootstrap5_List.length; i++) {
    if (empCodeStr == bootstrap5_List[i][0].toString()) {
      var bootstrap5Values = sheet_quiz_bootstrap5.getRange(i + 1, 1, 1, sheet_quiz_bootstrap5.getLastColumn()).getValues()[0];
      var bootstrap5Obj = {
        timestamp: bootstrap5Values[0].toString(),
        score: bootstrap5Values[1].toString(),
        empCode: bootstrap5Values[27].toString(),
        empEmail: bootstrap5Values[28].toString(),
      }
      quiz_bootstrap5_Array = [];
      quiz_bootstrap5_Array.push(bootstrap5Obj);
    }
  }

  for (var i = 0; i < php_List.length; i++) {
    if (empCodeStr == php_List[i][0].toString()) {
      var phpValues = sheet_quiz_PHP.getRange(i + 1, 1, 1, sheet_quiz_PHP.getLastColumn()).getValues()[0];
      var phpObj = {
        timestamp: phpValues[0].toString(),
        score: phpValues[1].toString(),
        empCode: phpValues[27].toString(),
        empEmail: phpValues[28].toString(),
      }
      quiz_php_Array = [];
      quiz_php_Array.push(phpObj);
    }
  }

  for (var i = 0; i < jQuery_List.length; i++) {
    if (empCodeStr == jQuery_List[i][0].toString()) {
      var jQueryValues = sheet_quiz_jQuery.getRange(i + 1, 1, 1, sheet_quiz_jQuery.getLastColumn()).getValues()[0];
      var jQueryObj = {
        timestamp: jQueryValues[0].toString(),
        score: jQueryValues[1].toString(),
        empCode: jQueryValues[27].toString(),
        empEmail: jQueryValues[28].toString(),
      }
      quiz_jQuery_Array = [];
      quiz_jQuery_Array.push(jQueryObj);
    }
  }

  for (var i = 0; i < cPlus_List.length; i++) {
    if (empCodeStr == cPlus_List[i][0].toString()) {
      var cPlusValues = sheet_quiz_CPlus.getRange(i + 1, 1, 1, sheet_quiz_CPlus.getLastColumn()).getValues()[0];
      var cPlusObj = {
        timestamp: cPlusValues[0].toString(),
        score: cPlusValues[1].toString(),
        empCode: cPlusValues[27].toString(),
        empEmail: cPlusValues[28].toString(),
      }
      quiz_cPlus_Array = [];
      quiz_cPlus_Array.push(cPlusObj);
    }
  }

  for (var i = 0; i < css_List.length; i++) {
    if (empCodeStr == css_List[i][0].toString()) {
      var cssValues = sheet_quiz_CSS.getRange(i + 1, 1, 1, sheet_quiz_CSS.getLastColumn()).getValues()[0];
      var cssObj = {
        timestamp: cssValues[0].toString(),
        score: cssValues[1].toString(),
        empCode: cssValues[27].toString(),
        empEmail: cssValues[28].toString(),
      }
      quiz_css_Array = [];
      quiz_css_Array.push(cssObj);
    }
  }

  for (var i = 0; i < javascript_List.length; i++) {
    if (empCodeStr == javascript_List[i][0].toString()) {
      var javascriptValues = sheet_quiz_Javascript.getRange(i + 1, 1, 1, sheet_quiz_Javascript.getLastColumn()).getValues()[0];
      var javascriptObj = {
        timestamp: javascriptValues[0].toString(),
        score: javascriptValues[1].toString(),
        empCode: javascriptValues[27].toString(),
        empEmail: javascriptValues[28].toString(),
      }
      quiz_javascript_Array = [];
      quiz_javascript_Array.push(javascriptObj);
    }
  }

  for (var i = 0; i < java_List.length; i++) {
    if (empCodeStr == java_List[i][0].toString()) {
      var javaValues = sheet_quiz_JAVA.getRange(i + 1, 1, 1, sheet_quiz_JAVA.getLastColumn()).getValues()[0];
      var javaObj = {
        timestamp: javaValues[0].toString(),
        score: javaValues[1].toString(),
        empCode: javaValues[27].toString(),
        empEmail: javaValues[28].toString(),
      }
      quiz_java_Array = [];
      quiz_java_Array.push(javaObj);
    }
  }

  for (var i = 0; i < html_List.length; i++) {
    if (empCodeStr == html_List[i][0].toString()) {
      var htmlValues = sheet_quiz_HTML.getRange(i + 1, 1, 1, sheet_quiz_HTML.getLastColumn()).getValues()[0];
      var htmlObj = {
        timestamp: htmlValues[0].toString(),
        score: htmlValues[1].toString(),
        empCode: htmlValues[42].toString(),
        empEmail: htmlValues[43].toString(),
      }
      quiz_html_Array = [];
      quiz_html_Array.push(htmlObj);
    }
  }

  return {
    success: true,
    message: 'getYourScoreBoard: success',
    quizXmlData: quiz_xml_Array,
    quizDotNetData: quiz_xml_Array,
    quizSqlData: quiz_sql_Array,
    quizTypescriptData: quiz_typescript_Array,
    quizResctData: quiz_react_Array,
    quizKotlinData: quiz_kotlin_Array,
    quizPostgressSqlData: quiz_postgressql_Array,
    quizPythonData: quiz_python_Array,
    quizBootstrap5Data: quiz_bootstrap5_Array,
    quizPhpData: quiz_php_Array,
    quizJQueryData: quiz_jQuery_Array,
    quizCPlusData: quiz_cPlus_Array,
    quizCssData: quiz_css_Array,
    quizJavascriptData: quiz_javascript_Array,
    quizJavaData: quiz_java_Array,
    quizHtmlData: quiz_html_Array,
  };
}

function getRecentGlobalQuiz() {
  var dotNet_List = sheet_quiz_dotNetCSharp.getRange("A:A").getValues();
  var xml_List = sheet_quiz_XML.getRange("A:A").getValues();
  var sql_List = sheet_quiz_SQL.getRange("A:A").getValues();
  var typescript_List = sheet_quiz_Typescript.getRange("A:A").getValues();
  var react_List = sheet_quiz_React.getRange("A:A").getValues();
  var kotlin_List = sheet_quiz_Kotlin.getRange("A:A").getValues();
  var postgreSql_List = sheet_quiz_PostgreSQL.getRange("A:A").getValues();
  var python_List = sheet_quiz_Python.getRange("A:A").getValues();
  var bootstrap5_List = sheet_quiz_bootstrap5.getRange("A:A").getValues();
  var php_List = sheet_quiz_PHP.getRange("A:A").getValues();
  var jQuery_List = sheet_quiz_jQuery.getRange("A:A").getValues();
  var cPlus_List = sheet_quiz_CPlus.getRange("A:A").getValues();
  var css_List = sheet_quiz_CSS.getRange("A:A").getValues();
  var javascript_List = sheet_quiz_Javascript.getRange("A:A").getValues();
  var java_List = sheet_quiz_JAVA.getRange("A:A").getValues();
  var html_List = sheet_quiz_HTML.getRange("A:A").getValues();

  let yesterday_Array = [];
  let today_Array = [];

  const today = convertStringToDate(new Date())
  const yesterday = convertStringToDate(new Date(new Date().getTime() - (24 * 60 * 60 * 1000)))

  for (var i = 1; i < html_List.length; i++) {
    var htmlValues = sheet_quiz_HTML.getRange(i + 1, 1, 1, sheet_quiz_HTML.getLastColumn()).getValues()[0];
    var empCode = htmlValues[42]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(html_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(htmlValues[0].toString()),
        score: htmlValues[1].toString(),
        quizName: "HTML"
      }
      today_Array.push(yesterdayObj)
    }
    else if (yesterday == convertStringToDate(html_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(htmlValues[0].toString()),
        score: htmlValues[1].toString(),
        quizName: "HTML"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (html_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < java_List.length; i++) {
    var rowValues = sheet_quiz_JAVA.getRange(i + 1, 1, 1, sheet_quiz_JAVA.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(java_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "JAVA"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(java_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "JAVA"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (java_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < javascript_List.length; i++) {
    var rowValues = sheet_quiz_Javascript.getRange(i + 1, 1, 1, sheet_quiz_Javascript.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(javascript_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "JavaScript"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(javascript_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "JavaScript"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (javascript_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < css_List.length; i++) {
    var rowValues = sheet_quiz_CSS.getRange(i + 1, 1, 1, sheet_quiz_CSS.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(css_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "CSS"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(css_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "CSS"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (css_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < cPlus_List.length; i++) {
    var rowValues = sheet_quiz_CPlus.getRange(i + 1, 1, 1, sheet_quiz_CPlus.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(cPlus_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "C, C++"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(cPlus_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "C, C++"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (cPlus_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < jQuery_List.length; i++) {
    var rowValues = sheet_quiz_jQuery.getRange(i + 1, 1, 1, sheet_quiz_jQuery.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(jQuery_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "jQuery"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(jQuery_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "jQuery"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (jQuery_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < php_List.length; i++) {
    var rowValues = sheet_quiz_PHP.getRange(i + 1, 1, 1, sheet_quiz_PHP.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(php_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "PHP"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(php_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "PHP"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (php_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < bootstrap5_List.length; i++) {
    var rowValues = sheet_quiz_bootstrap5.getRange(i + 1, 1, 1, sheet_quiz_bootstrap5.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(bootstrap5_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Bootstrap5"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(bootstrap5_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Bootstrap5"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (bootstrap5_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < python_List.length; i++) {
    var rowValues = sheet_quiz_Python.getRange(i + 1, 1, 1, sheet_quiz_Python.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(python_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Python"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(python_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Python"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (python_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < postgreSql_List.length; i++) {
    var rowValues = sheet_quiz_PostgreSQL.getRange(i + 1, 1, 1, sheet_quiz_PostgreSQL.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(postgreSql_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "PostgreSQL"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(postgreSql_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "PostgreSQL"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (postgreSql_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < kotlin_List.length; i++) {
    var rowValues = sheet_quiz_Kotlin.getRange(i + 1, 1, 1, sheet_quiz_Kotlin.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(kotlin_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Kotlin"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(kotlin_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Kotlin"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (kotlin_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < react_List.length; i++) {
    var rowValues = sheet_quiz_React.getRange(i + 1, 1, 1, sheet_quiz_React.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(react_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "React"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(react_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "React"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (react_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < typescript_List.length; i++) {
    var rowValues = sheet_quiz_Typescript.getRange(i + 1, 1, 1, sheet_quiz_Typescript.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(typescript_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Typescript"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(typescript_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "Typescript"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (typescript_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < sql_List.length; i++) {
    var rowValues = sheet_quiz_SQL.getRange(i + 1, 1, 1, sheet_quiz_SQL.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(sql_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "SQL"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(sql_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "SQL"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (sql_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < xml_List.length; i++) {
    var rowValues = sheet_quiz_XML.getRange(i + 1, 1, 1, sheet_quiz_XML.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(xml_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "XML"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(xml_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: "XML"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (xml_List[i][0].toString() == '') {
      break;
    }
  }

  for (var i = 1; i < dotNet_List.length; i++) {
    var rowValues = sheet_quiz_dotNetCSharp.getRange(i + 1, 1, 1, sheet_quiz_dotNetCSharp.getLastColumn()).getValues()[0];
    var empCode = rowValues[27]
    var userProfile = getUserProfile(empCode)

    if (today == convertStringToDate(dotNet_List[i][0].toString())) {
      var todayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: ".NET C#"
      }
      today_Array.push(todayObj)
    }
    else if (yesterday == convertStringToDate(dotNet_List[i][0].toString())) {
      var yesterdayObj = {
        ...userProfile,
        timestamp: convertStringToDate(rowValues[0].toString()),
        score: rowValues[1].toString(),
        quizName: ".NET C#"
      }
      yesterday_Array.push(yesterdayObj)
    }
    else if (dotNet_List[i][0].toString() == '') {
      break;
    }
  }

  return {
    success: true,
    message: 'getYourScoreBoard: success',
    recentYesterdayData: yesterday_Array,
    recentTodayData: today_Array,
  };
}


function getIndexPage() {
  return HtmlService.createHtmlOutputFromFile('index').getContent()
}

function getDataChart() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answer')
  console.info(data.getRange(1, 2, data.getLastRow(), 2).getValues())
  return data.getRange(1, 2, data.getLastRow(), 2).getValues()
}

function getDataTable() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answer')
  const table = data.getRange(1, 1, data.getLastRow(), data.getLastColumn()).getDisplayValues()
  console.log(table)
  return table
}
