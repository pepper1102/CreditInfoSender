/**
 * 設定項目
 */
const CONFIG = {
  GMAIL: {
    SOURCES: [
      { 
        NAME: '三井住友カード',
        FROM: 'statement@vpass.ne.jp',
        SUBJECT: 'ご利用のお知らせ【三井住友カード】',
        CYCLE_START: 1,   // 毎月1日開始
        CYCLE_END: 0      // 翌月0日=月末まで
      },
      { 
        NAME: 'ビューカード',
        FROM: 'viewcard@mail.viewsnet.jp',
        SUBJECT: '－確報版－ ビューカードご利用情報のお知らせ（本人会員利用）',
        CYCLE_START: 6,   // 毎月6日開始
        CYCLE_END: 5      // 翌月5日終了
      },
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
// 共通キャッシュ用
const mailCache = {};
// CONFIG 直後に追加してください
function getSource(cardName) {
  if (!CONFIG || !CONFIG.GMAIL || !Array.isArray(CONFIG.GMAIL.SOURCES)) {
    throw new Error('CONFIG.GMAIL.SOURCES が未定義です');
  }
  const src = CONFIG.GMAIL.SOURCES.find(s => s.NAME === cardName);
  if (!src) throw new Error(`未定義のカード: ${cardName}`);

  // 互換: CYCLE_START/CYCLE_END が無い場合は CYCLE_TYPE から補完
  if (typeof src.CYCLE_START === 'undefined' || typeof src.CYCLE_END === 'undefined') {
    if (src.CYCLE_TYPE === 'calendar') {
      // 毎月1日〜月末締め
      src.CYCLE_START = 1;
      src.CYCLE_END = 0; // 0 = 翌月0日（=前月末）
    } else if (src.CYCLE_TYPE === '5日締め') {
      // 6日開始〜翌月5日終了
      src.CYCLE_START = 6;
      src.CYCLE_END = 5;
    } else {
      throw new Error(`CYCLE_START/CYCLE_END 未設定（CYCLE_TYPE から補完不可）: ${cardName}`);
    }
  }
  return src;
}
// メイン
function sendDailyCardUsageToLINE() {

  const props = PropertiesService.getScriptProperties();
  const todayStr = formatDate(new Date());
  const results = CONFIG.GMAIL.SOURCES.map(src => {
    const cycleStart = getCycleStart(src.NAME, new Date());
    const endExclusive = new Date();
    return { name: src.NAME, total: sumMonthlyAmountFromGmail(buildQueryForCard(src.NAME, cycleStart, endExclusive), src.NAME) };
  });

  // メッセージ生成部を分離
  const message = buildCardUsageMessage(todayStr, results);
  sendLineMessageByPush(message);

  const saveObj = {};
  for (const r of results) saveObj[r.name] = r.total;
  props.setProperty(CONFIG.PROPS.LAST_TOTAL, JSON.stringify(saveObj));
  props.setProperty(CONFIG.PROPS.LAST_DATE, todayStr);
}


function buildQueryForCard(cardName, startInclusive, endInclusive) {
  const src = getSource(cardName);
  if (!src) throw new Error(`未定義のカード: ${cardName}`);

  const afterDate = addDays(startInclusive, -1); // 開始日前日
  const beforeDate = addDays(endInclusive, +1);  // 終了日翌日

  let subjectPart = '';
  if (cardName === 'ビューカード') {
    // 件名の表記ゆれに強い「トークン AND」検索
    // 例: subject:(ビューカード ご利用情報)
    subjectPart = 'subject:(ビューカード ご利用情報)';
  } else if (src.SUBJECT) {
    // 三井住友は従来の完全一致でもOK
    subjectPart = `subject:"${src.SUBJECT}"`;
  }

  const parts = [
    `from:${src.FROM}`,
    subjectPart,
    `after:${formatDate(afterDate)}`,
    `before:${formatDate(beforeDate)}`,
  ].filter(Boolean);

  return parts.join(' ');
}

// 締め日判定（利用日がサイクル内かどうか）
function isWithinBillingPeriod(usageDate, cardName, today) {
  if (!usageDate) return false;
  const start = getCycleStart(cardName, today);
  const end = getCycleEnd(cardName, today);
  return usageDate >= start && usageDate <= end;
}

// Gmail集計（キャッシュ対応）
function sumMonthlyAmountFromGmail(query, cardName, startInclusive, endInclusive) {
  const cacheKey = `${cardName}:${query}:${startInclusive ? formatDate(startInclusive) : ''}:${endInclusive ? formatDate(endInclusive) : ''}`;
  if (mailCache[cacheKey] !== undefined) return mailCache[cacheKey];

  let total = 0;
  let threads = GmailApp.search(query, 0, 500);

  // ビューカード 0件フォールバック（subject ゆるめ）
  if (threads.length === 0 && cardName === 'ビューカード' && startInclusive && endInclusive) {
    const src = getSource(cardName);
    const fallback = [
      `from:${src.FROM}`,
      `subject:ご利用情報`,
      `after:${formatDate(addDays(startInclusive, -1))}`,
      `before:${formatDate(addDays(endInclusive, +1))}`,
    ].join(' ');
    console.info('View fallback query:', fallback);
    threads = GmailApp.search(fallback, 0, 500);
  }

  // --- 三井住友カード：従来どおり加算（重複は基本起こりにくい前提） ---
  if (cardName !== 'ビューカード') {
    for (const th of threads) {
      for (const msg of th.getMessages()) {
        const raw = msg.getPlainBody() || stripHtml(msg.getBody() || '');
        const text = sanitizeText(raw);
        const usageDate = extractUsageDate(text, cardName);
        const basisDate = usageDate || msg.getDate();
        if (startInclusive && endInclusive && !inRangeInclusive(basisDate, startInclusive, endInclusive)) continue;

        const delta = sumAmountsMitsui(text);
        total += delta;
      }
    }
    mailCache[cacheKey] = total;
    return total;
  }


  // --- ビューカード：確報優先（重複排除）※キーは「日付＋金額」に一本化 ---
  const confirmed = new Map();   // hardKey(yyyyMMdd\n金額) -> amt
  const provisional = new Map(); // hardKey(yyyyMMdd\n金額) -> amt
   for (const th of threads) {
     for (const msg of th.getMessages()) {
       const subj = msg.getSubject() || '';
       const typ = classifyViewSubject(subj); // confirmed / provisional / unknown
       const raw = msg.getPlainBody() || stripHtml(msg.getBody() || '');
       const text = sanitizeText(raw);
       const usageDate = extractUsageDate(text, cardName);
      const basisDate = usageDate || msg.getDate();
       if (!basisDate) continue;
       if (startInclusive && endInclusive && !inRangeInclusive(basisDate, startInclusive, endInclusive)) continue;
       const amt = sumAmountsViewCard(text);
       if (!amt) continue;
       const ymd = toYmd(basisDate);
      const hardKey = `${ymd}\n${amt}`; // ← 日付＋金額のみで紐づけ
      if (typ === 'confirmed' || typ === 'unknown') {
        confirmed.set(hardKey, amt);
      } else if (typ === 'provisional') {
        if (confirmed.has(hardKey)) continue; // 同キー確報があれば速報は除外
        provisional.set(hardKey, amt);        // 上書きで最新を保持
      }
     }
   }
  // 合計：確報は全採用
  for (const v of confirmed.values()) total += v;
  // 合計：確報に吸収されていない速報だけ採用
  for (const [k, v] of provisional.entries()) {
    if (confirmed.has(k)) continue;
    total += v;
  }
   mailCache[cacheKey] = total;
   return total;
 }









// --- 差し替え: 三井住友カード 金額抽出（円/JPY、ラベル優先） ---
function sumAmountsMitsui(text) {
  if (!text) return 0;
  let sum = 0, m;

  // 1) ラベル明示（最優先）
  // 例: ご利用金額：1,234円 / 金額：¥1,234 / 12,000.00 JPY 等
  const reStrict = /(?:ご利用金額|金額)\s*[：:]\s*(?:[¥￥]\s*)?([\d,]+)(?:\s*円|(?:\.\d{2})?\s*JPY)?/gi;
  while ((m = reStrict.exec(text)) !== null) {
    sum += parseInt((m[1] || '0').replace(/,/g, ''), 10);
  }
  if (sum > 0) return sum;

  // 2) フォールバック：利用行に出る金額（円）
  const reFallbackYen = /[^\n]*(?:ご利用|利用)[^\n]*?([¥￥]?\s*[\d,]+)\s*円/gi;
  while ((m = reFallbackYen.exec(text)) !== null) {
    const v = (m[1] || '').replace(/[¥￥,\s]/g, '');
    if (v) sum += parseInt(v, 10);
  }
  if (sum > 0) return sum;

  // 3) さらにフォールバック：JPY（小数点は切り捨て扱い）
  const reJPY = /([0-9][\d,]*)\s*(?:JPY)/gi;
  while ((m = reJPY.exec(text)) !== null) {
    sum += parseInt(m[1].replace(/,/g, ''), 10);
  }
  return sum;
}

// --- 差し替え: ビューカード 金額抽出（「税込」などカッコ書き対応） ---
function sumAmountsViewCard(text) {
  if (!text) return 0;
  let sum = 0, m;

  // 例: 利用金額：2,000円（税込） / ご利用金額：3,000円（内税等）
  // カッコ内コメントは無視して金額だけ拾う
  const re = /(?:利用金額|ご利用金額)\s*[：:]\s*(?:[¥￥]\s*)?([\d,]+)\s*円(?:（.*?）|\(.*?\))?/gi;
  while ((m = re.exec(text)) !== null) {
    sum += parseInt((m[1] || '0').replace(/,/g, ''), 10);
  }

  // ビューカードは基本これで十分ですが、必要ならフォールバック（稀）
  if (sum === 0) {
    const reLoose = /[^\n]*(?:利用金額|ご利用金額)[^\n]*?([¥￥]?\s*[\d,]+)\s*円/gi;
    while ((m = reLoose.exec(text)) !== null) {
      const v = (m[1] || '').replace(/[¥￥,\s]/g, '');
      if (v) sum += parseInt(v, 10);
    }
  }
  return sum;
}


// 前月サイクルの合計
function getCardPrevMonthTotal(cardName) {
  const today = new Date();
  const src = CONFIG.GMAIL.SOURCES.find(s => s.NAME === cardName);

  // 「前サイクル」の基準日を作る（今月から1ヶ月前）
  const baseDate = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
  const start = getCycleStart(cardName, baseDate);
  const end = getCycleEnd(cardName, baseDate);

  const total = sumMonthlyAmountFromGmail(buildQueryForCard(cardName, start, end), cardName);


// 昨日まで（サイクル内）をプロパティで増分キャッシュ
const yesterdayTotal = getMtdYesterdayTotalCached(cardName, start, end);


  // 月ラベルは終了日の月を採用
  const monthLabel = end.getMonth() + 1; // JSは0始まりなので+1

  return { total, diff: total - yesterdayTotal, month: monthLabel };
}

// 当月サイクルの合計
function getCardCurrentMonthTotal(cardName) {
  const today = new Date();

  const start = getCycleStart(cardName, today);
  const end = getCycleEnd(cardName, today);

  const total = sumMonthlyAmountFromGmail(buildQueryForCard(cardName, start, end), cardName);

// 昨日まで（サイクル内）をプロパティで増分キャッシュ
const yesterdayTotal = getMtdYesterdayTotalCached(cardName, start, end);


  // 月ラベルは開始日の月を採用（「今サイクル」として扱うため）
  const monthLabel = start.getMonth() + 1;

  return { total, diff: total - yesterdayTotal, month: monthLabel };
}

// 文字列整形
function formatJPY(n){ return Number(n).toLocaleString('ja-JP'); }

// LINE送信
function sendLineMessageByPush(message){
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(CONFIG.PROPS.CHANNEL_ACCESS_TOKEN);
  const to = props.getProperty(CONFIG.PROPS.LINE_TO_USER_ID);
  if(!token||!to) throw new Error('LINE情報が未設定');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
    method:'post',
    headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},
    payload:JSON.stringify({to,messages:[{type:'text',text:message}]}),
    muteHttpExceptions:true
  });
  if(res.getResponseCode()<200||res.getResponseCode()>=300) throw new Error(`LINE送信失敗: ${res.getContentText()}`);
}



// メッセージ生成専用関数
function buildCardUsageMessage(todayStr, results) {
  let message = `${todayStr}時点でのカード利用額は 以下の通りです\n`;
  if (new Date().getDate() <= 5) {
    const prevMitsui = getCardPrevMonthTotal('三井住友カード');
    const prevView = getCardPrevMonthTotal('ビューカード');
    message += `${prevMitsui.month}月の利用額：¥${formatJPY(prevMitsui.total + prevView.total)}\n`;
    message += `（三井住友カード：¥${formatJPY(prevMitsui.total)}（前日より+¥${formatJPY(prevMitsui.diff)}）\n`;
    message += `ビューカード：¥${formatJPY(prevView.total)}（前日より+¥${formatJPY(prevView.diff)}））\n\n`;
    const curMitsui = getCardCurrentMonthTotal('三井住友カード');
    message += `${curMitsui.month}月の利用額：¥${formatJPY(curMitsui.total)}\n`;
    message += `（三井住友カード：¥${formatJPY(curMitsui.total)}（前日より+¥${formatJPY(curMitsui.diff)}））`;
  } else {
    let totalAll = 0, diffAll = 0;
    for (const r of results) {
      const cycleStart = getCycleStart(r.name, new Date());
      const cycleEnd = getCycleEnd(r.name, new Date());
      const yesterdayTotal = getMtdYesterdayTotalCached(r.name, cycleStart, cycleEnd);
      const diff = r.total - yesterdayTotal;
      message += `${r.name}：¥${formatJPY(r.total)}（前日より${diff >= 0 ? '+' : '-'}¥${formatJPY(Math.abs(diff))}）\n`;
      totalAll += r.total; diffAll += diff;
    }
    message += `総合計：¥${formatJPY(totalAll)}（前日より${diffAll >= 0 ? '+' : '-'}¥${formatJPY(Math.abs(diffAll))}）`;
  }
  return message;
}

function classifyViewSubject(subj) {
  subj = (subj || '').trim();
  if (/確報版/.test(subj)) return 'confirmed';
  if (/速報版/.test(subj)) return 'provisional';
  return 'unknown';
}

function extractMerchant(text) {
  // ビュー：主に「利用先」、三井住友：場合により「ご利用加盟店名」
  const m =
    text.match(/(?:利用先|ご利用加盟店名)\s*[：:]\s*(.+)/);
  return m ? m[1].trim() : '';
}

function toYmd(d) {
  return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), d.getDate()), CONFIG.TZ, 'yyyyMMdd');
}

/**
 * MTD（サイクル開始～前日(yEnd)まで）の合計を PropertiesService で増分キャッシュ。
 * 前回保存した「前日」から新しい「前日」までの差分だけを Gmail から再集計する。
 */
function getMtdYesterdayTotalCached(cardName, cycleStart, cycleEnd) {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.tryLock(3000); // 競合実行の保護（任意）

  try {
    const today = new Date();
    const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    // 前日とサイクル終了のうち、早い方を yEnd とする（サイクル外へ出ないように）
    const yEnd = (yesterday < cycleEnd) ? yesterday : cycleEnd;

    // 前日がサイクル開始より前なら 0 固定
    if (compareYmd(yEnd, cycleStart) < 0) {
      return 0;
    }

    const base = `mtd:${cardName}:${toYmd(cycleStart)}-${toYmd(cycleEnd)}`;
    const lastDateKey  = `${base}:lastDate`;
    const lastTotalKey = `${base}:lastTotal`;

    const lastDateStr  = props.getProperty(lastDateKey);
    const lastTotalStr = props.getProperty(lastTotalKey);

    // 初回：yEnd までをフル集計して保存
    if (!lastDateStr || !lastTotalStr) {
      const total = sumMonthlyAmountFromGmail(
        buildQueryForCard(cardName, cycleStart, yEnd),
        cardName,
        cycleStart,
        yEnd
      );
      props.setProperty(lastDateKey,  toYmd(yEnd));
      props.setProperty(lastTotalKey, String(total));
      return total;
    }

    const lastDate = parseYmd(lastDateStr);
    let total = Number(lastTotalStr) || 0;

    // 同じ yEnd まで計算済みなら再集計不要
    if (compareYmd(lastDate, yEnd) >= 0) {
      return total;
    }

    // 差分： (lastDate + 1日) ～ yEnd を追加集計
    const from = addDays(lastDate, 1);
    const delta = sumMonthlyAmountFromGmail(
      buildQueryForCard(cardName, from, yEnd),
      cardName,
      from,
      yEnd
    );
    total += delta;

    props.setProperty(lastDateKey,  toYmd(yEnd));
    props.setProperty(lastTotalKey, String(total));
    return total;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

