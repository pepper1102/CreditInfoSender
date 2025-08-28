/**
 * 設定項目
 */
const CONFIG = {
  GMAIL: {
    SOURCES: [
      { NAME: '三井住友カード', FROM: 'statement@vpass.ne.jp', SUBJECT: 'ご利用のお知らせ【三井住友カード】' },
      { NAME: 'ビューカード', FROM: 'viewcard@mail.viewsnet.jp', SUBJECT: '－確報版－ ビューカードご利用情報のお知らせ（本人会員利用）' }, // 件名条件不要なら空文字
    ],
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

  // Gmail 検索範囲（当月のみ）
  const afterStr = Utilities.formatDate(firstDay, CONFIG.TZ, 'yyyy/MM/dd');
  const beforeStr = Utilities.formatDate(nextMonthFirst, CONFIG.TZ, 'yyyy/MM/dd');

  // 各送信元を集計
  const results = [];
for (const src of CONFIG.GMAIL.SOURCES) {
  const query = [
    `from:${src.FROM}`,
    src.SUBJECT ? `subject:"${src.SUBJECT}"` : '',
    `after:${afterStr}`,
    `before:${beforeStr}`,
  ].filter(Boolean).join(' ');

  const monthTotal = sumMonthlyAmountFromGmail(query, src.NAME);
  results.push({ name: src.NAME, total: monthTotal });
}

  // 前回保存値
  const lastSavedDate = props.getProperty(CONFIG.PROPS.LAST_DATE) || '';
  const lastTotalRaw = props.getProperty(CONFIG.PROPS.LAST_TOTAL);
  const lastTotal = lastTotalRaw ? JSON.parse(lastTotalRaw) : {}; // {sourceName: amount}

  const isDifferentMonth =
    lastSavedDate &&
    Utilities.formatDate(new Date(lastSavedDate), CONFIG.TZ, 'yyyyMM') !==
      Utilities.formatDate(now, CONFIG.TZ, 'yyyyMM');

  // メッセージ組み立て
  let message = `${todayStr}時点でのカード利用額は\n`;
  let totalAll = 0;
  let diffAll = 0;

  for (const r of results) {
    const prevTotal = isDifferentMonth ? 0 : (lastTotal[r.name] || 0);
    const diff = r.total - prevTotal;
    const diffSign = diff >= 0 ? '+' : '-';
    message += `${r.name}：¥${formatJPY(r.total)}（前日より${diffSign}¥${formatJPY(Math.abs(diff))}）\n`;

    totalAll += r.total;
    diffAll += diff;
  }

  // 総合計を追加
  const diffAllSign = diffAll >= 0 ? '+' : '-';
  message += `総合計：¥${formatJPY(totalAll)}（前日より${diffAllSign}¥${formatJPY(Math.abs(diffAll))}）`;

  // LINE 送信
  sendLineMessageByPush(message);

  // 状態保存（送信元ごとにJSONで）
  const saveObj = {};
  for (const r of results) saveObj[r.name] = r.total;
  props.setProperty(CONFIG.PROPS.LAST_TOTAL, JSON.stringify(saveObj));
  props.setProperty(CONFIG.PROPS.LAST_DATE, todayStr);
}

/**
 * Gmail から当月メールを集計して累計額を返す
 */
function sumMonthlyAmountFromGmail(query, sourceName) {
  let total = 0;

  const threads = GmailApp.search(query, 0, 500);
  for (const th of threads) {
    const msgs = th.getMessages();
    for (const msg of msgs) {
      const text = sanitizeText(msg.getPlainBody() || '');
      if (sourceName === '三井住友カード') {
        total += sumAmountsMitsui(text);
      } else if (sourceName === 'ビューカード') {
        total += sumAmountsViewCard(text);
      }
    }
  }
  return total;
}
/**
 * 三井住友カード用：「ご利用内容」以降のテキストから「◯◯円」を抽出
 */
function sumAmountsMitsui(text) {
  if (!text) return 0;
  const marker = 'ご利用内容';
  const idx = text.indexOf(marker);
  const target = idx >= 0 ? text.slice(idx + marker.length) : text;

  const re = /(\d{1,3}(?:,\d{3})*|\d+)\s*円/g;
  let sum = 0, m;
  while ((m = re.exec(target)) !== null) {
    const n = Number(String(m[1]).replace(/,/g, ''));
    if (!Number.isNaN(n)) sum += n;
  }
  return sum;
}

/**
 * ビューカード用：「利用金額」の行から「◯◯円」を抽出
 */
function sumAmountsViewCard(text) {
  if (!text) return 0;
  const re = /利用金額\s*[:：]\s*([\d,]+)円/g;
  let sum = 0, m;
  while ((m = re.exec(text)) !== null) {
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

  if (!token) throw new Error('CHANNEL_ACCESS_TOKEN が未設定です。ScriptPropertiesに値を設定してください');
  if (!to) throw new Error('LINE_TO_USER_ID が未設定です。ScriptPropertiesに値を設定してください');

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
