/**
 * 設定項目
 */
const CONFIG = {
  GMAIL: {
    SOURCES: [
      { NAME: '三井住友カード', FROM: 'statement@vpass.ne.jp', SUBJECT: 'ご利用のお知らせ【三井住友カード】' },
      { NAME: 'ビューカード', FROM: 'viewcard@mail.viewsnet.jp', SUBJECT: '－確報版－ ビューカードご利用情報のお知らせ（本人会員利用）' },
    ],
  },
  PROPS: {
    LAST_TOTAL: 'lastTotalAmount',
    LAST_DATE: 'lastSavedDate',
    CHANNEL_ACCESS_TOKEN: 'CHANNEL_ACCESS_TOKEN',
    LINE_TO_USER_ID: 'LINE_TO_USER_ID',
  },
  TZ: 'Asia/Tokyo',
};

/**
 * メイン：毎日23:00に実行想定
 */
function sendDailyCardUsageToLINE() {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TZ, 'yyyy/MM/dd');

  const lastSavedDate = props.getProperty(CONFIG.PROPS.LAST_DATE) || '';
  const lastTotalRaw = props.getProperty(CONFIG.PROPS.LAST_TOTAL);
  const lastTotal = lastTotalRaw ? JSON.parse(lastTotalRaw) : {};
  const isDifferentMonth =
    lastSavedDate &&
    Utilities.formatDate(new Date(lastSavedDate), CONFIG.TZ, 'yyyyMM') !==
      Utilities.formatDate(now, CONFIG.TZ, 'yyyyMM');

  // カードごとの集計
  const results = [];

  for (const src of CONFIG.GMAIL.SOURCES) {
    let afterStr, beforeStr;

    if (src.NAME === 'ビューカード') {
      // ビューカード: 締め日5日 → 6日〜翌月5日
      const year = now.getFullYear();
      const month = now.getMonth(); // 0-based
      let start, end;

      if (now.getDate() <= 5) {
        start = new Date(year, month - 1, 6);
        end   = new Date(year, month, 6);
      } else {
        start = new Date(year, month, 6);
        end   = new Date(year, month + 1, 6);
      }

      afterStr  = Utilities.formatDate(start, CONFIG.TZ, 'yyyy/MM/dd');
      beforeStr = Utilities.formatDate(end,   CONFIG.TZ, 'yyyy/MM/dd');

} else if (src.NAME === '三井住友カード') {
  // 三井住友カード: 月初〜月末
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-based
  const firstDay = new Date(y, m, 1);
  const lastDay  = new Date(y, m + 1, 0); // 今月末日

  // Gmail検索は "after:前月末" ～ "before:翌月1日"
  const prevMonthEnd = new Date(y, m, 0);    // 前月末
  const nextMonthFirst = new Date(y, m + 1, 1);

  afterStr  = Utilities.formatDate(prevMonthEnd, CONFIG.TZ, 'yyyy/MM/dd');
  beforeStr = Utilities.formatDate(nextMonthFirst, CONFIG.TZ, 'yyyy/MM/dd');
}


    const query = [
      `from:${src.FROM}`,
      src.SUBJECT ? `subject:"${src.SUBJECT}"` : '',
      `after:${afterStr}`,
      `before:${beforeStr}`,
    ].filter(Boolean).join(' ');

    const monthTotal = sumMonthlyAmountFromGmail(query, src.NAME);
    results.push({ name: src.NAME, total: monthTotal });
  }

  // -------- メッセージ組み立て --------
  let message = `${todayStr}時点でのカード利用額は 以下の通りです\n`;

  if (now.getDate() <= 5) {
    // 1日〜5日：前月分 + 当月分（三井のみ）

    // 前月分
    const prevMitsui = getCardPrevMonthTotal('三井住友カード', lastTotal, isDifferentMonth);
    const prevView   = getCardPrevMonthTotal('ビューカード', lastTotal, isDifferentMonth);
    const prevSum = prevMitsui.total + prevView.total;

    message += `${prevMitsui.month}月の利用額：¥${formatJPY(prevSum)}\n`;
    message += `（三井住友カード：¥${formatJPY(prevMitsui.total)}（前日より+¥${formatJPY(prevMitsui.diff)}）\n`;
    message += `ビューカード：¥${formatJPY(prevView.total)}（前日より+¥${formatJPY(prevView.diff)}））\n\n`;

    // 当月分（三井のみ）
    const curMitsui = getCardCurrentMonthTotal('三井住友カード', results, lastTotal, isDifferentMonth);
    message += `${curMitsui.month}月の利用額：¥${formatJPY(curMitsui.total)}\n`;
    message += `（三井住友カード：¥${formatJPY(curMitsui.total)}（前日より+¥${formatJPY(curMitsui.diff)}））`;

  } else {
    // 6日以降：通常通り合算
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

    const diffAllSign = diffAll >= 0 ? '+' : '-';
    message += `総合計：¥${formatJPY(totalAll)}（前日より${diffAllSign}¥${formatJPY(Math.abs(diffAll))}）`;
  }

  // LINE送信
  sendLineMessageByPush(message);

  // 状態保存
  const saveObj = {};
  for (const r of results) saveObj[r.name] = r.total;
  props.setProperty(CONFIG.PROPS.LAST_TOTAL, JSON.stringify(saveObj));
  props.setProperty(CONFIG.PROPS.LAST_DATE, todayStr);
}

/**
 * カードごとの前月集計取得（ビューカードもGmailから集計）
 */
function getCardPrevMonthTotal(cardName, lastTotal, isDifferentMonth) {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth(); // 0-based
  let total = 0, afterStr, beforeStr;

if (cardName === '三井住友カード') {
  // 前月1日〜前月末
  const prevMonthFirst = new Date(y, m - 1, 1);   // 前月1日
  const prevMonthEnd   = new Date(y, m, 0);       // 前月末
  const thisMonthFirst = new Date(y, m, 1);       // 今月1日

  // Gmail検索は "after:前月1日-1日" ～ "before:今月1日"
  const afterDate  = new Date(prevMonthFirst.getFullYear(), prevMonthFirst.getMonth(), prevMonthFirst.getDate() - 1);
  afterStr  = Utilities.formatDate(afterDate, CONFIG.TZ, 'yyyy/MM/dd');
  beforeStr = Utilities.formatDate(thisMonthFirst, CONFIG.TZ, 'yyyy/MM/dd');
}else if (cardName === 'ビューカード') {
  // 前月6日〜今月5日
  const prevMonth6th = new Date(y, m - 1, 6);  // 前月6日
  const thisMonth6th = new Date(y, m, 6);      // 今月6日

  // Gmail検索は "after:前月5日 before:今月6日"
  const afterDate = new Date(prevMonth6th.getFullYear(), prevMonth6th.getMonth(), prevMonth6th.getDate() - 1);
  afterStr  = Utilities.formatDate(afterDate, CONFIG.TZ, 'yyyy/MM/dd');
  beforeStr = Utilities.formatDate(thisMonth6th, CONFIG.TZ, 'yyyy/MM/dd');
}

  const src = CONFIG.GMAIL.SOURCES.find(s => s.NAME === cardName);
  const query = [
    `from:${src.FROM}`,
    src.SUBJECT ? `subject:"${src.SUBJECT}"` : '',
    `after:${afterStr}`,
    `before:${beforeStr}`,
  ].filter(Boolean).join(' ');

  total = sumMonthlyAmountFromGmail(query, cardName);

  return { total: total, diff: 0, month: today.getMonth() };
}

/**
 * カードごとの当月集計（三井のみ）
 */
function getCardCurrentMonthTotal(cardName, results, lastTotal, isDifferentMonth) {
  const today = new Date();
  const month = today.getMonth() + 1;
  const card = results.find(r => r.name === cardName) || { total: 0 };
  const prevTotal = isDifferentMonth ? 0 : (lastTotal[cardName] || 0);
  const diff = card.total - prevTotal;
  return { total: card.total, diff: diff, month: month };
}

/* 既存の sumMonthlyAmountFromGmail, sumAmountsMitsui, sumAmountsViewCard, sendLineMessageByPush, sanitizeText, formatJPY, setupTriggerAt2300, resetSavedState はそのまま */

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

 // ◯◯円 または ◯◯.00 JPY に対応
  const re = /(\d{1,3}(?:,\d{3})*|\d+)(?:\s*円|\.\d{2}\s*JPY)/g;
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

/**
 * 手動リセット
 */
function resetSavedState() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CONFIG.PROPS.LAST_TOTAL);
  props.deleteProperty(CONFIG.PROPS.LAST_DATE);
}



