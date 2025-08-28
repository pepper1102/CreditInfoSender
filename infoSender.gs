/**
 * 設定項目
 */
const CONFIG = {
  GMAIL: {
    FROM: 'statement@vpass.ne.jp',
    SUBJECT: 'ご利用のお知らせ【三井住友カード】',
  },
  PROPS: {
    LAST_TOTAL: 'lastTotalAmount',
    LAST_DATE: 'lastSavedDate',
    CHANNEL_ACCESS_TOKEN: 'CHANNEL_ACCESS_TOKEN', // スクリプトプロパティに保存
    LINE_TO_USER_ID: 'LINE_TO_USER_ID',           // スクリプトプロパティに保存
  },
  TZ: 'Asia/Tokyo',
};

/**
 * メイン：毎日23:00に実行想定
 */
function sendDailyCardUsageToLINE() {
  const props = PropertiesService.getScriptProperties();

  // 今日・当月境界（日付はJST）
  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TZ, 'yyyy/MM/dd');
  const y = Number(Utilities.formatDate(now, CONFIG.TZ, 'yyyy'));
  const m = Number(Utilities.formatDate(now, CONFIG.TZ, 'MM')) - 1; // 0-based
  const firstDay = new Date(Date.UTC(y, m, 1));
  const nextMonthFirst = new Date(Date.UTC(y, m + 1, 1));

  // Gmail 検索クエリ（当月のみ）
  const afterStr = Utilities.formatDate(firstDay, CONFIG.TZ, 'yyyy/MM/dd');
  const beforeStr = Utilities.formatDate(nextMonthFirst, CONFIG.TZ, 'yyyy/MM/dd');
  const query = [
    `from:${CONFIG.GMAIL.FROM}`,
    `subject:"${CONFIG.GMAIL.SUBJECT}"`,
    `after:${afterStr}`,
    `before:${beforeStr}`,
  ].join(' ');

  // 当月累計額の算出
  const monthTotal = sumMonthlyAmountFromGmail(query);

  // 前日累計の取得と月替わりリセット
  const lastSavedDate = props.getProperty(CONFIG.PROPS.LAST_DATE) || '';
  const lastTotalRaw = props.getProperty(CONFIG.PROPS.LAST_TOTAL);
  const lastTotal = lastTotalRaw ? Number(lastTotalRaw) : 0;

  const isDifferentMonth =
    lastSavedDate &&
    Utilities.formatDate(new Date(lastSavedDate), CONFIG.TZ, 'yyyyMM') !==
      Utilities.formatDate(now, CONFIG.TZ, 'yyyyMM');

  const prevTotalForDiff = isDifferentMonth ? 0 : lastTotal;

  // 差額計算（当日累計 - 前日累計）
  const diff = monthTotal - prevTotalForDiff;

  // メッセージ生成（所定フォーマット）
  const dateForMsg = Utilities.formatDate(now, CONFIG.TZ, 'yyyy/MM/dd');
  const diffSign = diff >= 0 ? '+' : '-';
  const diffAbs = Math.abs(diff);
  const message =
    `${dateForMsg}時点でのカード利用額は\n` +
    `¥${formatJPY(monthTotal)}（前日より${diffSign}¥${formatJPY(diffAbs)}）です。`;

  // LINE 送信（Messaging API Push）
  sendLineMessageByPush(message);

  // 状態保存
  props.setProperty(CONFIG.PROPS.LAST_TOTAL, String(monthTotal));
  props.setProperty(CONFIG.PROPS.LAST_DATE, todayStr);
}

/**
 * Gmail から当月メールを集計して累計額を返す
 */
function sumMonthlyAmountFromGmail(query) {
  let total = 0;

  const threads = GmailApp.search(query, 0, 500);
  for (const th of threads) {
    const msgs = th.getMessages();
    for (const msg of msgs) {
      const text = sanitizeText(msg.getPlainBody() || '');
      //const html = sanitizeText(stripHtml(msg.getBody() || ''));
      total += sumAmountsAfterMarker(text);
      //total += sumAmountsAfterMarker(html);
    }
  }
  return total;
}

/**
 * 「ご利用内容」以降のテキストから「◯◯円」を抽出して合算
 */
function sumAmountsAfterMarker(text) {
  if (!text) return 0;
  const marker = 'ご利用内容';
  const idx = text.indexOf(marker);
  const target = idx >= 0 ? text.slice(idx + marker.length) : text;

  const re = /(\d{1,3}(?:,\d{3})*|\d+)\s*円/g;
  let m;
  let sum = 0;
  while ((m = re.exec(target)) !== null) {
    const n = Number(String(m[1]).replace(/,/g, ''));
    if (!Number.isNaN(n)) sum += n;
  }
  return sum;
}

/**
 * LINE Messaging API に Push 送信
 */
function sendLineMessageByPush(message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(CONFIG.PROPS.CHANNEL_ACCESS_TOKEN);
  const to = props.getProperty(CONFIG.PROPS.LINE_TO_USER_ID);

  if (!token) throw new Error('CHANNEL_ACCESS_TOKEN が未設定です。ScriptPropatiesに値を設定してください');
  if (!to) throw new Error('LINE_TO_USER_ID が未設定です。ScriptPropatiesに値を設定してください');

  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to,
    messages: [{ type: 'text', text: message }],
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`LINE Push 送信失敗: ${code} ${res.getContentText()}`);
  }
}

/**
 * 文字列整形
 */
function stripHtml(html) {
  return html.replace(/<[^>]*>/g, ' ');
}
function sanitizeText(s) {
  return s
    .replace(/\u00A0/g, ' ')
    .replace(/\r/g, '\n')
    .replace(/\t/g, ' ')
    .replace(/ +/g, ' ')
    .trim();
}
function formatJPY(n) {
  try {
    return Number(n).toLocaleString('ja-JP');
  } catch (e) {
    return String(n);
  }
}

/**
 * 初回セットアップ用：毎日23:00トリガー登録
 */
function setupTriggerAt2300() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'sendDailyCardUsageToLINE') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('sendDailyCardUsageToLINE')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();
}

/**
 * 手動リセット
 */
function resetSavedState() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CONFIG.PROPS.LAST_TOTAL);
  props.deleteProperty(CONFIG.PROPS.LAST_DATE);
}
