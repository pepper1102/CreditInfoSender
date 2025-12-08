/**
 * ログレベル定義
 */
const LogLevel = {
    ERROR: 0,
    WARN: 1,
    INFO: 2,
    DEBUG: 3
};

/**
 * 現在のログレベル設定
 * 本番環境では INFO、開発環境では DEBUG を推奨
 */
const CURRENT_LOG_LEVEL = LogLevel.INFO;

/**
 * エラーログを出力（常に出力）
 * @param {string} message - ログメッセージ
 * @param {*} data - 追加データ（オプション）
 */
function logError(message, data) {
    if (CURRENT_LOG_LEVEL >= LogLevel.ERROR) {
        if (data !== undefined) {
            console.error(`[ERROR] ${message}`, data);
        } else {
            console.error(`[ERROR] ${message}`);
        }
    }
}

/**
 * 警告ログを出力
 * @param {string} message - ログメッセージ
 * @param {*} data - 追加データ（オプション）
 */
function logWarn(message, data) {
    if (CURRENT_LOG_LEVEL >= LogLevel.WARN) {
        if (data !== undefined) {
            console.warn(`[WARN] ${message}`, data);
        } else {
            console.warn(`[WARN] ${message}`);
        }
    }
}

/**
 * 情報ログを出力（本番環境での標準的なログ）
 * @param {string} message - ログメッセージ
 * @param {*} data - 追加データ（オプション）
 */
function logInfo(message, data) {
    if (CURRENT_LOG_LEVEL >= LogLevel.INFO) {
        if (data !== undefined) {
            console.log(`[INFO] ${message}`, data);
        } else {
            console.log(`[INFO] ${message}`);
        }
    }
}

/**
 * デバッグログを出力（開発環境のみ）
 * @param {string} message - ログメッセージ
 * @param {*} data - 追加データ（オプション）
 */
function logDebug(message, data) {
    if (CURRENT_LOG_LEVEL >= LogLevel.DEBUG) {
        if (data !== undefined) {
            console.log(`[DEBUG] ${message}`, data);
        } else {
            console.log(`[DEBUG] ${message}`);
        }
    }
}

/**
 * カード集計結果のログ出力
 * @param {string} cardName - カード名
 * @param {Date} cycleStart - サイクル開始日
 * @param {Date} cycleEnd - サイクル終了日
 * @param {Date} endExclusive - 集計終了日
 * @param {number} total - 集計金額
 */
function logCardSummary(cardName, cycleStart, cycleEnd, endExclusive, total) {
    logDebug(`集計結果 - ${cardName}`, {
        cycleStart: toYmd(cycleStart),
        cycleEnd: toYmd(cycleEnd),
        endExclusive: toYmd(endExclusive),
        total
    });
}

/**
 * プロパティ保存のログ出力
 * @param {Object} saveObj - 保存するオブジェクト
 * @param {string} todayStr - 日付文字列
 */
function logSaveConfirmation(saveObj, todayStr) {
    logDebug('Saving to properties', {
        LAST_TOTAL: saveObj,
        LAST_DATE: todayStr,
        LAST_TOTAL_KEY: CONFIG.PROPS.LAST_TOTAL,
        LAST_DATE_KEY: CONFIG.PROPS.LAST_DATE
    });

    // 保存後の確認（DEBUGレベル）
    const props = PropertiesService.getScriptProperties();
    logDebug('Saved values', {
        LAST_TOTAL: props.getProperty(CONFIG.PROPS.LAST_TOTAL),
        LAST_DATE: props.getProperty(CONFIG.PROPS.LAST_DATE)
    });
}
