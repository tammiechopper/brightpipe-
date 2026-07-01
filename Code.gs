const SHEET_ID = ;
function getSheet(name) {
  return SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(name);
}function testSheet() {
  const ss =
    SpreadsheetApp.openById(SHEET_ID);

  Logger.log(ss.getName());
}
/**
 * 明管 Bright Pipe MVP
 * Sewer Status Event Bridge
 * https://github.com/tammiechopper/brightpipe-
 */


// ====================================================
// Step 1: 拉取水利署即時水位 OpenData
// ====================================================
function fetchWaterLevel() {
  const url = 'https://opendata.wra.gov.tw/api/v2/73c4c3de-4045-4765-abeb-89f9f9cd5ff0?sort=_importdate%20asc&format=JSON';

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Bright-Pipe-MVP)'
      }
    });

    const code = response.getResponseCode();
    Logger.log('HTTP 狀態碼: ' + code);

    if (code !== 200) {
      Logger.log('❌ 拉取失敗');
      return null;
    }

    const json = JSON.parse(response.getContentText());
    const data = json.responseData || json;

    Logger.log('✅ 成功拉到 ' + data.length + ' 筆水位資料');
    return data;
  } catch (error) {
    Logger.log('❌ 錯誤: ' + error.toString());
    return null;
  }
}


// ====================================================
// Step 2: 分析全台水位站分布、找高水位站、找雲林相關
// ====================================================
function analyzeStations() {
  const data = fetchWaterLevel();
  if (!data) {
    Logger.log('沒拉到資料,跳出');
    return;
  }

  // 篩出有效資料
  const valid = data.filter(d =>
    d.waterlevel &&
    d.waterlevel !== '' &&
    !isNaN(parseFloat(d.waterlevel))
  );

  Logger.log('═══════════════════════');
  Logger.log('📊 全台水位站分析');
  Logger.log('═══════════════════════');
  Logger.log('總站數: ' + data.length);
  Logger.log('有效站數: ' + valid.length);

  // 站號前綴分組
  const byPrefix = {};
  valid.forEach(d => {
    const prefix = d.stationid.substring(0, 4);
    if (!byPrefix[prefix]) byPrefix[prefix] = [];
    byPrefix[prefix].push(d);
  });

  Logger.log('');
  Logger.log('--- 各區域站數分布(站號前 4 碼) ---');
  Object.keys(byPrefix).sort().forEach(p => {
    Logger.log(p + ': ' + byPrefix[p].length + ' 站');
  });

  // 此刻水位最高 10 站
  Logger.log('');
  Logger.log('--- 此刻水位最高 10 站 ---');
  const sorted = valid.slice().sort((a, b) =>
    parseFloat(b.waterlevel) - parseFloat(a.waterlevel)
  );
  sorted.slice(0, 10).forEach((d, i) => {
    Logger.log((i + 1) + '. ' + d.stationid + ' → ' + d.waterlevel + ' m  (' + d.datetime + ')');
  });

  // 雲林相關(站號 1140 / 1310 開頭)
  Logger.log('');
  Logger.log('--- 雲林相關(站號 1140/1310 開頭) ---');
  const yunlin = valid.filter(d =>
    d.stationid.startsWith('1140') ||
    d.stationid.startsWith('1310')
  );
  Logger.log('找到 ' + yunlin.length + ' 個可能在雲林的站');
  yunlin.slice(0, 10).forEach(d => {
    Logger.log('  ' + d.stationid + ' → ' + d.waterlevel + ' m');
  });
}// ====================================================
// Step 3: Web App 入口(doGet) — 顯示設定介面
// ====================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('明管 Bright Pipe · 設定介面')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}function getDashboardData() {

  const data = fetchWaterLevel();

  if (!data) {
    return {
      total: 0,
      warning: 0,
      critical: 0,
      stations: []
    };
  }

  const valid = data.filter(d =>
    d.waterlevel &&
    !isNaN(parseFloat(d.waterlevel))
  );

  const warning = valid.filter(
    d => parseFloat(d.waterlevel) >= 1
  );

  const critical = valid.filter(
    d => parseFloat(d.waterlevel) >= 1.5
  );

  return {
    total: valid.length,
    warning: warning.length,
    critical: critical.length,
    stations: valid.slice(0,20)
  };
}
function sendTestAlert() {
  const sheet = getSheet('notifications');
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const alertId = 'A' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd') + '-001';

  const targets = [
    { channel: 'Telegram', recipient: '四湖里里長',   role: '地方' },
    { channel: 'Email',    recipient: '下水道科值班', role: '主管機關' },
    { channel: 'Telegram', recipient: '清潔隊分隊',   role: '現場處置' },
    { channel: 'Telegram', recipient: '在地防災志工', role: '疏散協助' }
  ];

  const receivers = [];
  Logger.log('TARGETS 數量: ' + TARGETS.length);   // 
   targets.forEach(function (t, i) {
    const notifId = 'N' + Utilities.formatDate(now, 'Asia/Taipei', 'HHmmss') + '-' + (i + 1);
    sheet.appendRow([
      notifId, alertId, t.channel, t.recipient, t.role,
      '', 'SENT', ts, 'PENDING', ''
    ]);
    receivers.push({ name: t.recipient, status: '已送達' });
  });

  Logger.log('已寫入 ' + targets.length + ' 筆到 notifications');
  return { receivers: receivers, alertId: alertId };
}
function markAck() {
  const sheet = getSheet('notifications');
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  // 找第一筆還在 PENDING 的,把它改成已收到
  const data = sheet.getDataRange().getValues();
  // 欄位: notifId|alertId|channel|recipient|role|messageId|sentStatus|sentAt|ackStatus|ackAt
  for (let r = 1; r < data.length; r++) {   // r=1 跳過標題列
    if (data[r][8] === 'PENDING') {
      sheet.getRange(r + 1, 9).setValue('ACK 已收到');  // 第9欄 ackStatus
      sheet.getRange(r + 1, 10).setValue(ts);           // 第10欄 ackAt
      Logger.log('已回寫抵達確認: ' + data[r][3] + ' @ ' + ts);
      return { recipient: data[r][3], ackAt: ts };
    }
  }
  Logger.log('沒有待確認的項目');
  return null;
}
// ===== Telegram 設定 =====
const TG_TOKEN = '8883827760:AAGGHfzP3J_AMZ3bB35TbNbkKEAgj0J2Uss';   // ← 換新 token 貼這,不要貼給別人

const TARGETS = [
  { channel:'Telegram', recipient:'四湖里里長',   role:'地方',     chatId:'8958874271' },
  { channel:'Telegram', recipient:'下水道科值班', role:'主管機關', chatId:'8958874271' },
  { channel:'Telegram', recipient:'清潔隊分隊',   role:'現場處置', chatId:'8958874271' },
  { channel:'Telegram', recipient:'在地防災志工', role:'疏散協助', chatId:'8958874271' }
];

// 真正打 Telegram
function sendTelegram(chatId, text, notifId) {
  const url = 'https://api.telegram.org/bot' + TG_TOKEN + '/sendMessage';
  const payload = {
    chat_id: chatId,
    text: text,
    reply_markup: {
      inline_keyboard: [[
        { text: '✅ 我已收到', callback_data: 'ack_' + notifId }
      ]]
    }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('TG 回應: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}
// 真推播 + 寫 notifications
function sendRealAlert() {
  const sheet = getSheet('notifications');
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const alertId = 'A' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd') + '-001';

  Logger.log('開始推播,TARGETS 數量: ' + TARGETS.length);

  const receivers = [];
  TARGETS.forEach(function (t, i) {
    const notifId = 'N' + Utilities.formatDate(now, 'Asia/Taipei', 'HHmmss') + '-' + (i + 1);
    const msg =
      '🚨 明管示警\n' +
      '四湖05站 水位1.75m(+0.63)\n' +
      '恐溢出中山東路\n' +
      '對象:' + t.recipient + '(' + t.role + ')\n' +
      '請確認是否收到';

    let sentStatus = 'FAILED';
    let messageId = '';
    const res = sendTelegram(t.chatId, msg, notifId);
    if (res.ok) {
      sentStatus = 'SENT';
      messageId = String(res.result.message_id);
    }

    sheet.appendRow([
      notifId, alertId, t.channel, t.recipient, t.role,
      messageId, sentStatus, ts, 'PENDING', ''
    ]);
    receivers.push({ name: t.recipient, status: sentStatus });
  });

  Logger.log('已推播 ' + TARGETS.length + ' 筆');
  return { receivers: receivers, alertId: alertId };
}
function doPost(e) {
  try {
    Logger.log('★ doPost 被觸發了,收到內容: ' + e.postData.contents);
    const update = JSON.parse(e.postData.contents);

    // 只處理按鈕事件
    if (update.callback_query) {
      const cq = update.callback_query;
      const data = cq.data;                 // 例如 ack_N170512-1
      const chatId = cq.message.chat.id;

      if (data.indexOf('ack_') === 0) {
        const notifId = data.substring(4);   // 去掉 ack_ 前綴
        ackByNotifId(notifId);

        // 回覆 Telegram:按鈕不轉圈、跳個小提示
        UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_TOKEN + '/answerCallbackQuery', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            callback_query_id: cq.id,
            text: '✅ 已記錄您的抵達確認'
          }),
          muteHttpExceptions: true
        });
      }
    }
  } catch (err) {
    Logger.log('doPost 錯誤: ' + err);
  }
  return ContentService.createTextOutput('ok');
}

// 依 notifId 找到那列,回寫 ack
function ackByNotifId(notifId) {
  const sheet = getSheet('notifications');
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const data = sheet.getDataRange().getValues();
  Logger.log('收到 notifId: [' + notifId + ']');   // ← 看按鈕傳來什麼
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === notifId) {
      sheet.getRange(r + 1, 9).setValue('ACK 已收到');
      sheet.getRange(r + 1, 10).setValue(ts);
      Logger.log('✅ 找到並回寫: ' + data[r][3]);
      return;
    }
  }
  Logger.log('❌ 找不到這個 notifId,表裡沒有');   // ← 對不上會印這行
}
