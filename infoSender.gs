
/**
 * カード利用額を集計し、LINEへ送信・保存するメイン関数
 * @returns {void}
 */
function sendDailyCardUsageToLINE() {

  const props = PropertiesService.getScriptProperties();
  const todayStr = formatDate(new Date());
  const results = CONFIG.GMAIL.SOURCES.map(src => {
    const cycleStart = getCycleStart(src.NAME, new Date());
    const cycleEnd = getCycleEnd(src.NAME, new Date());
    const endExclusive = new Date();
    const query = buildQueryForCard(src.NAME, cycleStart, endExclusive);
    const total = sumMonthlyAmountFromGmail(query, src.NAME, cycleStart, endExclusive);
    console.log(`集計結果 - ${src.NAME}:`, {
      cycleStart: toYmd(cycleStart),
      cycleEnd: toYmd(cycleEnd),
      endExclusive: toYmd(endExclusive),
      total
    });
    return { name: src.NAME, total, cycleStart, cycleEnd };
  });

  const message = buildCardUsageMessage(todayStr, results);
  sendLineMessageByPush(message);

  // 結果をプロパティに保存（カードごとに意味が分かる形で保存）
  const saveObj = {};
  for (const r of results) {
    saveObj[r.name] = { total: r.total, cycleStart: toYmd(r.cycleStart), cycleEnd: toYmd(r.cycleEnd) };
  }
  console.log('Saving to properties:', {
    LAST_TOTAL: saveObj,
    LAST_DATE: todayStr,
    LAST_TOTAL_KEY: CONFIG.PROPS.LAST_TOTAL,
    LAST_DATE_KEY: CONFIG.PROPS.LAST_DATE
  });
  props.setProperty(CONFIG.PROPS.LAST_TOTAL, JSON.stringify(saveObj));
  props.setProperty(CONFIG.PROPS.LAST_DATE, todayStr);
  
  // 保存後の確認
  console.log('Saved values:', {
    LAST_TOTAL: props.getProperty(CONFIG.PROPS.LAST_TOTAL),
    LAST_DATE: props.getProperty(CONFIG.PROPS.LAST_DATE)
  });

}
/**
 * カード利用額の集計結果からLINE送信用メッセージを生成
 * @param {string} todayStr - 日付文字列
 * @param {Array<{name:string,total:number}>} results - 集計結果配列
 * @returns {string} 送信メッセージ
 */
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
      try {
      const yesterdayTotal = getMtdYesterdayTotalCached(r.name, cycleStart, cycleEnd);
      const diff = r.total - yesterdayTotal;
      if (yesterdayTotal === 0 && toYmd(cycleStart) === formatDate(new Date())) {
        message += `${r.name}：¥${formatJPY(r.total)}（今日からの新サイクル）\n`;
      } else {
        message += `${r.name}：¥${formatJPY(r.total)}（前日より${diff >= 0 ? '+' : '-'}¥${formatJPY(Math.abs(diff))}）\n`;
      }
      totalAll += r.total; diffAll += diff;
      } catch (e) {
        console.error(`Error processing ${r.name}:`, e);
        message += `${r.name}：¥${formatJPY(r.total)}（前日比較不可）\n`;
        totalAll += r.total;
      }

    }
    message += `総合計：¥${formatJPY(totalAll)}（前日より${diffAll >= 0 ? '+' : '-'}¥${formatJPY(Math.abs(diffAll))}）`;
  }
  return message;
}
/**
 * LINEへメッセージをPush送信する
 * @param {string} message - 送信するテキスト
 * @throws {Error} LINE情報未設定・送信失敗時
 * @returns {void}
 */
function sendLineMessageByPush(message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(CONFIG.PROPS.CHANNEL_ACCESS_TOKEN);
  const to = props.getProperty(CONFIG.PROPS.LINE_TO_USER_ID);
  if (!token || !to) throw new Error('LINE情報が未設定');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ to, messages: [{ type: 'text', text: message }] }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error(`LINE送信失敗: ${res.getContentText()}`);
}
/**
 * Gmailからカード利用額を集計（キャッシュ・ビューカード重複排除対応）
 * @param {string} query - Gmail検索クエリ
 * @param {string} cardName - カード名
 * @param {Date} startInclusive - 集計開始日
 * @param {Date} endInclusive - 集計終了日
 * @returns {number} 利用額合計
 */
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
  if (cardName === '三井住友カード') {
    for (const th of threads) {
      for (const msg of th.getMessages()) {
        const raw = msg.getPlainBody() || stripHtml(msg.getBody() || '');
        const text = sanitizeText(raw);
        const usageDate = extractUsageDate(text, cardName);
        const basisDate = usageDate || msg.getDate();
        if (startInclusive && endInclusive && !inRangeInclusive(basisDate, startInclusive, endInclusive)) {
          console.log(`${cardName}の利用を除外:`, {
            利用日: toYmd(basisDate),
            サイクル開始: toYmd(startInclusive),
            サイクル終了: toYmd(endInclusive)
          });
          continue;
        }
        //console.log(text)
        const delta = sumAmountsMitsui(text);
        total += delta;
      }
    }
    mailCache[cacheKey] = total;
    return total;
  }


  // --- ビューカード：確報優先（重複排除）※キーは「日付＋金額」に一本化 ---
  // confirmed: 確報（またはunknown）メールの金額を格納するMap
  // provisional: 速報メールの金額を格納するMap
  // hardKeyは「日付+金額」で一意に紐づけ
  const confirmed = new Map();   // hardKey(yyyyMMdd\n金額) -> amt
  const provisional = new Map(); // hardKey(yyyyMMdd\n金額) -> amt
  for (const th of threads) {
    for (const msg of th.getMessages()) {
      const subj = msg.getSubject() || '';
      // 件名から「確報/速報/unknown」を判定
      const typ = classifyViewSubject(subj); // confirmed / provisional / unknown
      // メール本文を整形
      const raw = msg.getPlainBody() || stripHtml(msg.getBody() || '');
      const text = sanitizeText(raw);
      // 利用日を抽出（なければメール日付）
      const usageDate = extractUsageDate(text, cardName);
      const basisDate = usageDate || msg.getDate();
      if (!basisDate) continue; // 日付がなければスキップ
      // サイクル範囲外は除外
      if (startInclusive && endInclusive && !inRangeInclusive(basisDate, startInclusive, endInclusive)) continue;
      // 金額抽出
      const amt = sumAmountsViewCard(text);
      if (!amt) continue; // 金額がなければスキップ
      // 日付＋金額で一意キー生成
      const ymd = toYmd(basisDate);
      const hardKey = `${ymd}\n${amt}`; // ← 日付＋金額のみで紐づけ
      if (typ === 'confirmed' || typ === 'unknown') {
        // 確報またはunknownはconfirmedに格納
        confirmed.set(hardKey, amt);
      } else if (typ === 'provisional') {
        // 速報は、同じキーの確報があれば除外
        if (confirmed.has(hardKey)) continue;
        // 速報は上書きで最新を保持
        provisional.set(hardKey, amt);
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
/**
 * カード名・期間からGmail検索クエリを生成
 * @param {string} cardName - カード名
 * @param {Date} startInclusive - 開始日
 * @param {Date} endInclusive - 終了日
 * @returns {string} Gmail検索クエリ
 */
function buildQueryForCard(cardName, startInclusive, endInclusive) {
  const src = getSource(cardName);
  if (!src) throw new Error(`未定義のカード: ${cardName}`);

  // メール受信日のマージンを広げる（利用日とメール受信日のずれを考慮）
  const afterDate = addDays(startInclusive, -2); // 開始日の2日前
  const beforeDate = addDays(endInclusive, +2);  // 終了日の2日後

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


/**
 * 三井住友カードのメール本文から金額を抽出
 * @param {string} text - メール本文
 * @returns {number} 抽出した金額合計
 */
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

/**
 * ビューカードのメール本文から金額を抽出（カッコ書き対応）
 * @param {string} text - メール本文
 * @returns {number} 抽出した金額合計
 */
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


/**
 * 前月サイクルの利用額合計・差分・月ラベルを取得
 * @param {string} cardName - カード名
 * @returns {{total:number, diff:number, month:number}} 集計結果
 */
function getCardPrevMonthTotal(cardName) {
  const today = new Date();
  const src = CONFIG.GMAIL.SOURCES.find(s => s.NAME === cardName);
  //console.log(src);

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

/**
 * 当月サイクルの利用額合計・差分・月ラベルを取得
 * @param {string} cardName - カード名
 * @returns {{total:number, diff:number, month:number}} 集計結果
 */
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

/**
 * 金額を日本語カンマ区切り文字列に変換
 * @param {number} n - 金額
 * @returns {string} 整形済み文字列
 */
function formatJPY(n) { return Number(n).toLocaleString('ja-JP'); }


/**
 * ビューカードメール件名から「確報/速報/unknown」を判定
 * @param {string} subj - 件名
 * @returns {string} 判定結果（confirmed/provisional/unknown）
 */
function classifyViewSubject(subj) {
  subj = (subj || '').trim();
  if (/確報版/.test(subj)) return 'confirmed';
  if (/速報版/.test(subj)) return 'provisional';
  return 'unknown';
}

/**
 * Date型をyyyyMMdd形式の文字列に変換
 * @param {Date} d - 日付
 * @returns {string} 変換後文字列
 */
function toYmd(d) {
  return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), d.getDate()), CONFIG.TZ, 'yyyyMMdd');
}

/**
 * サイクル開始～前日までの合計をPropertiesServiceで増分キャッシュ
 * @param {string} cardName - カード名
 * @param {Date} cycleStart - サイクル開始日
 * @param {Date} cycleEnd - サイクル終了日
 * @returns {number} 合計金額
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

    // 正規化キー：カードごとにサイクルを YYYYMM で一意にする（cycleStart の年月を使用）
    const cycleLabel = `${cycleStart.getFullYear()}${String(cycleStart.getMonth() + 1).padStart(2, '0')}`;
    const base = `mtd:${cardName}:${cycleLabel}`;
    const lastDateKey = `${base}:lastDate`;
    const lastTotalKey = `${base}:lastTotal`;

    // 前日がサイクル開始より前の場合
    if (compareYmd(yEnd, cycleStart) < 0) {
      console.log(`${cardName}: サイクル開始日(${toYmd(cycleStart)})のため、前日比較なし`);
      // サイクル開始日の場合は、現在の値を初期値として保存
      if (compareYmd(new Date(), cycleStart) === 0) {
        const today = new Date();
        const query = buildQueryForCard(cardName, cycleStart, today);
        console.log(`サイクル開始日の初期値を集計 - ${cardName}:`, {
          query,
          start: toYmd(cycleStart),
          end: toYmd(today)
        });
        const currentTotal = sumMonthlyAmountFromGmail(query, cardName, cycleStart, today);
        console.log(`サイクル開始日の集計結果 - ${cardName}:`, currentTotal);
        props.setProperty(lastDateKey, toYmd(new Date()));
        props.setProperty(lastTotalKey, String(currentTotal));
      }
      return 0;
    }



    const lastDateStr = props.getProperty(lastDateKey);
    const lastTotalStr = props.getProperty(lastTotalKey);

    // 許容する保存キー（カードごとに当月と前月のみ保持）
    const curCycleStart = getCycleStart(cardName, new Date());
    const prevCycleStart = new Date(curCycleStart.getFullYear(), curCycleStart.getMonth() - 1, curCycleStart.getDate());
    const allowedLabels = [
      `${curCycleStart.getFullYear()}${String(curCycleStart.getMonth() + 1).padStart(2, '0')}`,
      `${prevCycleStart.getFullYear()}${String(prevCycleStart.getMonth() + 1).padStart(2, '0')}`
    ];
    const allowedPrefixes = allowedLabels.map(l => `mtd:${cardName}:${l}`);

    // 初回：yEnd までをフル集計して保存
    if (!lastDateStr || !lastTotalStr) {
      const total = sumMonthlyAmountFromGmail(
        buildQueryForCard(cardName, cycleStart, yEnd),
        cardName,
        cycleStart,
        yEnd
      );
      props.setProperty(lastDateKey, toYmd(yEnd));
      props.setProperty(lastTotalKey, String(total));

      // クリーンアップ：当カードの mtd:... は当月・前月のみ残す
      const all = props.getProperties();
      for (const k of Object.keys(all)) {
        if (k.startsWith(`mtd:${cardName}:`) && !allowedPrefixes.some(p => k.startsWith(p))) {
          props.deleteProperty(k);
        }
      }

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

    props.setProperty(lastDateKey, toYmd(yEnd));
    props.setProperty(lastTotalKey, String(total));

    // クリーンアップ：当カードの mtd:... は当月・前月のみ残す
    const all2 = props.getProperties();
    for (const k of Object.keys(all2)) {
      if (k.startsWith(`mtd:${cardName}:`) && !allowedPrefixes.some(p => k.startsWith(p))) {
        props.deleteProperty(k);
      }
    }

    return total;
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

