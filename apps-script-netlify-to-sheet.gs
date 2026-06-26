function formatHeader() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = [
    '送出時間', '中文姓名', '英文習慣用名', '手機號碼', 'LINE帳號', 'Email',
    '生理性別', '報名方案', '匯款帳號資訊', '推薦人', '是否上過課',
    '曾上過的課程', '同意健康聲明', '同意退費條款', '確認時間地點'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#b8923f');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  var subId = body.id;
  if (subId) {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'sub_' + subId;
    if (cache.get(cacheKey)) {
      return ContentService.createTextOutput(JSON.stringify({status: 'duplicate_ignored'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    cache.put(cacheKey, '1', 300);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    formatHeader();
  }

  var data = body.data || {};

  sheet.appendRow([
    new Date(),
    data['中文姓名'] || '',
    data['英文習慣用名'] || '',
    data['手機號碼'] || '',
    data['LINE帳號'] || '',
    data['Email'] || '',
    data['gender'] || '',
    data['plan'] || '',
    data['匯款帳號資訊'] || '',
    data['推薦人'] || '',
    data['hasExperience'] || '',
    data['曾上過的課程'] || '',
    data['同意健康聲明'] ? '是' : '否',
    data['同意退費條款'] ? '是' : '否',
    data['確認時間地點'] ? '是' : '否'
  ]);

  if (data['Email']) {
    sendConfirmationEmail(data);
  }

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendConfirmationEmail(data) {
  var name = data['中文姓名'] || '同學';
  var plan = data['plan'] || '';

  var body =
    name + '，你好：\n\n' +
    '感謝你報名「愛的源頭・家族系統排列工作坊 with Praful Saracino」！\n\n' +
    '我們已經收到你的報名資料：\n' +
    '報名方案：' + plan + '\n\n' +
    '【下一步】請主動透過 LINE 或 WhatsApp 與我們的工作人員聯繫，確認名額並完成匯款，報名才算正式成立：\n' +
    'Jennie 侯君妮　LINE：jenniecnho　WhatsApp：+886-988-817-609\n' +
    'Connie 蔡亞亭　LINE：connie216　WhatsApp：+886-988-255-964\n\n' +
    '工作坊資訊：\n' +
    '日期：9月2日（週三）－9月6日（週日），每日 10:00–18:00\n' +
    '地點：兆基文教大樓・台北市松山區南京東路四段120巷11號5樓\n\n' +
    '期待與你相見！\n' +
    '愛的源頭・家族系統排列工作坊';

  MailApp.sendEmail({
    to: data['Email'],
    subject: '🌿 報名確認｜愛的源頭・家族系統排列工作坊',
    body: body
  });
}
