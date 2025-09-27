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
/**
 * カード名から設定情報（SOURCES配列の要素）を取得する
 * @param {string} cardName - カード名
 * @returns {Object} カード設定情報
 * @throws {Error} 未定義カードや設定不足の場合
 */
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

/**
 * 日付を指定フォーマット（yyyy/MM/dd）で文字列化
 * @param {Date} d - 日付オブジェクト
 * @returns {string} フォーマット済み日付文字列
 */
function formatDate(d) {
  return Utilities.formatDate(d, CONFIG.TZ, 'yyyy/MM/dd');
}

function getCycleStart(cardName, refDate) {
  const src = getSource(cardName); // 既存の getSource を利用
  const y = refDate.getFullYear();
  const m = refDate.getMonth();    // 0-based
  const d = refDate.getDate();

    // ref 日が開始日より前なら「前月の開始日」、それ以外は「当月の開始日」
    return (d < src.CYCLE_START)
      ? new Date(y, m - 1, src.CYCLE_START)
      : new Date(y, m, src.CYCLE_START);
}

function getCycleEnd(cardName, refDate) {
  const src = getSource(cardName);
  const start = getCycleStart(cardName, refDate);

  if (src.CYCLE_END === 0) {
    // 月末締め → 開始月の末日
    return new Date(start.getFullYear(), start.getMonth() + 1, 0);
  }

  // 月跨ぎかどうか：例）6開始・5締め → 跨ぐ（<=）
  const crossesMonth = src.CYCLE_END <= src.CYCLE_START;
  const endMonthOffset = crossesMonth ? 1 : 0;

  return new Date(
    start.getFullYear(),
    start.getMonth() + endMonthOffset,
    src.CYCLE_END
  );
}

/**
 * 指定日付に日数を加算（または減算）した新しいDateを返す
 * @param {Date} d - 基準日
 * @param {number} n - 加算/減算する日数
 * @returns {Date} 加算後の日付
 */
function addDays(d, n) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() + n);
  return x;
}
/**
 * 日付dがstart～endの範囲（両端含む）に入っているか判定
 * @param {Date} d - 判定する日付
 * @param {Date} start - 範囲開始
 * @param {Date} end - 範囲終了
 * @returns {boolean} 範囲内ならtrue
 */
function inRangeInclusive(d, start, end) {
  if (!(d instanceof Date)) return false;
  // 時刻の影響を除くため「日付のみ」で比較
  var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return x >= start && x <= end;
}

/**
 * 指定日付が属するサイクルの開始・終了日を返す
 * @param {string} cardName - カード名
 * @param {Date} [refDate=new Date()] - 基準日
 * @returns {{start: Date, end: Date}} サイクル範囲
 */
function getCycleRange(cardName, refDate = new Date()) {
  const start = getCycleStart(cardName, refDate);
  const end   = getCycleEnd(cardName, refDate);
  return { start, end };
}

/**
 * 前サイクルの開始・終了日を返す
 * @param {string} cardName - カード名
 * @param {Date} [refDate=new Date()] - 基準日
 * @returns {{start: Date, end: Date}} 前サイクル範囲
 */
function getPrevCycleRange(cardName, refDate = new Date()) {
  const currentStart = getCycleStart(cardName, refDate);
  const prevRef = addDays(currentStart, -1);
  const start = getCycleStart(cardName, prevRef);
  const end   = getCycleEnd(cardName, prevRef);
  return { start, end };
}

/**
 * 全角数字や記号を半角に変換する（ゆるやか正規化）
 * @param {string} s - 入力文字列
 * @returns {string} 正規化済み文字列
 */
function normalizeDigits(s) {
  try {
    return s.normalize('NFKC');
  } catch (e) {
    return s
      .replace(/[０-９]/g, function(ch){ return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
      .replace(/[－―ー‐]/g, '-')
      .replace(/[，]/g, ',')
      .replace(/[：]/g, ':')
      .replace(/[￥]/g, '¥');
  }
}

/**
 * HTMLタグを除去し、style/scriptタグも安全化
 * @param {string} s - 入力HTML文字列
 * @returns {string} タグ除去済み文字列
 */
function stripHtml(s) {
  return (s || '')
    .replace(/<style[\s\S]*?<\x2Fstyle>/gi, ' ')
    .replace(/<script[\s\S]*?<\x2Fscript>/gi, ' ')
    .replace(/<[^>]+>/g, ' ');
}

/**
 * テキストを正規化し、改行・空白を整理
 * @param {string} s - 入力文字列
 * @returns {string} 整形済み文字列
 */
function sanitizeText(s) {
  return normalizeDigits(
    (s || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\r/g, '\n')
      .replace(/\t/g, ' ')
      .replace(/ +/g, ' ')
      .replace(/\n{2,}/g, '\n')
      .trim()
  );
}


/**
 * メール本文から利用日を抽出（ラベル優先・フォールバックあり）
 * @param {string} text - メール本文
 * @param {string} cardName - カード名
 * @returns {Date|null} 抽出した日付（失敗時はnull）
 */
function extractUsageDate(text, cardName) {
  // 1) 正規化（HTML除去 → 全角→半角 → 空白整理）
  var s = sanitizeText(stripHtml(text || ''));

  // 2) ラベル優先：
  //    「ご利用日時 / 利用日時 / ご利用日 / 利用日」の後に
  //    2025/09/12, 2025-9-2, 2025年9月02日のいずれか
  var reLabeled =
    /(?:ご利用(?:日時|日)|利用(?:日時|日))\s*[：:]\s*(\d{4})(?:\/|-|年)(\d{1,2})(?:\/|-|月)(\d{1,2})(?:日)?/i;

  var m = s.match(reLabeled);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

  // 3) フォールバック：
  //    行内の最初の「yyyy/mm/dd | yyyy-mm-dd | yyyy年m月d日（時刻付きOK）」を拾う
  var reLoose =
    /(\d{4})(?:\/|-|年)(\d{1,2})(?:\/|-|月)(\d{1,2})(?:日)?(?:\s+\d{1,2}:\d{2})?/;

  m = s.match(reLoose);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

  return null;
}
/**
 * yyyyMMdd形式の文字列をDate型に変換
 * @param {string} ymd - 日付文字列（例: 20250927）
 * @returns {Date} Dateオブジェクト
 */
function parseYmd(ymd) {
  const y = Number(ymd.substring(0, 4));
  const m = Number(ymd.substring(4, 6)) - 1;
  const d = Number(ymd.substring(6, 8));
  return new Date(y, m, d);
}

/**
 * 日付（時分秒無視）の大小比較
 * @param {Date} a - 比較対象A
 * @param {Date} b - 比較対象B
 * @returns {number} a<b:-1, a=b:0, a>b:1
 */
function compareYmd(a, b) {
  const ax = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const bx = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  if (ax.getTime() === bx.getTime()) return 0;
  return ax.getTime() < bx.getTime() ? -1 : 1;
}

