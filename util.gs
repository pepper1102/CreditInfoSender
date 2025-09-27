/**
 * 日付ユーティリティ
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

// まだなければヘルパーも（存在していれば不要）
function addDays(d, n) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() + n);
  return x;
}
// 日付 d が start～end の「両端を含む」かを判定
function inRangeInclusive(d, start, end) {
  if (!(d instanceof Date)) return false;
  // 時刻の影響を除くため「日付のみ」で比較
  var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return x >= start && x <= end;
}

// refDate の属する「現サイクル」の開始/終了を返す
function getCycleRange(cardName, refDate = new Date()) {
  const start = getCycleStart(cardName, refDate);
  const end   = getCycleEnd(cardName, refDate);
  return { start, end };
}

// 「現サイクル開始日の前日」を基準に、前サイクルの開始/終了を返す
function getPrevCycleRange(cardName, refDate = new Date()) {
  const currentStart = getCycleStart(cardName, refDate);
  const prevRef = addDays(currentStart, -1);
  const start = getCycleStart(cardName, prevRef);
  const end   = getCycleEnd(cardName, prevRef);
  return { start, end };
}

// 全角→半角のゆるやか正規化
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

// HTML 除去（</style> / </script> の / を \x2F で安全化）
function stripHtml(s) {
  return (s || '')
    .replace(/<style[\s\S]*?<\x2Fstyle>/gi, ' ')
    .replace(/<script[\s\S]*?<\x2Fscript>/gi, ' ')
    .replace(/<[^>]+>/g, ' ');
}

// テキスト整形（normalize → 改行・空白の整理）
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


// 日付抽出（ラベル優先 → ゆるめフォールバック）
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
// 'yyyyMMdd' → Date
function parseYmd(ymd) {
  const y = Number(ymd.substring(0, 4));
  const m = Number(ymd.substring(4, 6)) - 1;
  const d = Number(ymd.substring(6, 8));
  return new Date(y, m, d);
}

// 日付（時分秒無視）の比較: a<b:-1, a=b:0, a>b:1
function compareYmd(a, b) {
  const ax = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  const bx = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  if (ax.getTime() === bx.getTime()) return 0;
  return ax.getTime() < bx.getTime() ? -1 : 1;
}

