/**
 * ユニットテスト - サイクル計算
 */
function testCycleCalculation() {
    logInfo('=== サイクル計算テスト開始 ===');

    // 三井住友カード: 1日開始、月末締め
    const mitsui_2025_01_15 = new Date(2025, 0, 15);
    const mitsui_start = getCycleStart('三井住友カード', mitsui_2025_01_15);
    const mitsui_end = getCycleEnd('三井住友カード', mitsui_2025_01_15);

    assertEqual(toYmd(mitsui_start), '20250101', '三井住友開始日');
    assertEqual(toYmd(mitsui_end), '20250131', '三井住友終了日');

    // ビューカード: 6日開始、翌月5日締め
    const view_2025_01_10 = new Date(2025, 0, 10);
    const view_start = getCycleStart('ビューカード', view_2025_01_10);
    const view_end = getCycleEnd('ビューカード', view_2025_01_10);

    assertEqual(toYmd(view_start), '20250106', 'ビューカード開始日');
    assertEqual(toYmd(view_end), '20250205', 'ビューカード終了日');

    // 境界値テスト: ビューカードの5日（締め日当日）
    const view_2025_02_05 = new Date(2025, 1, 5);
    const view_start_boundary = getCycleStart('ビューカード', view_2025_02_05);
    const view_end_boundary = getCycleEnd('ビューカード', view_2025_02_05);

    assertEqual(toYmd(view_start_boundary), '20250106', 'ビューカード境界値開始日');
    assertEqual(toYmd(view_end_boundary), '20250205', 'ビューカード境界値終了日');

    logInfo('✅ サイクル計算テスト成功');
}

/**
 * ユニットテスト - 金額抽出（三井住友）
 */
function testSumAmountsMitsui() {
    logInfo('=== 三井住友金額抽出テスト開始 ===');

    const testCases = [
        {
            text: 'ご利用金額：1,234円',
            expected: 1234,
            desc: '基本パターン'
        },
        {
            text: '金額：¥5,678円',
            expected: 5678,
            desc: '円記号付き'
        },
        {
            text: 'ご利用 ¥12,000円',
            expected: 12000,
            desc: 'フォールバック'
        },
        {
            text: 'ご利用金額：10,000.00 JPY',
            expected: 10000,
            desc: 'JPY表記'
        },
        {
            text: '利用金額なし',
            expected: 0,
            desc: '金額なし'
        }
    ];

    testCases.forEach(tc => {
        const result = sumAmountsMitsui(tc.text);
        assertEqual(result, tc.expected, `三井住友金額抽出: ${tc.desc}`);
    });

    logInfo('✅ 三井住友金額抽出テスト成功');
}

/**
 * ユニットテスト - 金額抽出（ビューカード）
 */
function testSumAmountsViewCard() {
    logInfo('=== ビューカード金額抽出テスト開始 ===');

    const testCases = [
        {
            text: 'ご利用金額：2,000円（税込）',
            expected: 2000,
            desc: 'カッコ書きあり'
        },
        {
            text: '利用金額：3,500円',
            expected: 3500,
            desc: '基本パターン'
        },
        {
            text: 'ご利用金額：¥1,234円',
            expected: 1234,
            desc: '円記号付き'
        },
        {
            text: '金額の記載なし',
            expected: 0,
            desc: '金額なし'
        }
    ];

    testCases.forEach(tc => {
        const result = sumAmountsViewCard(tc.text);
        assertEqual(result, tc.expected, `ビューカード金額抽出: ${tc.desc}`);
    });

    logInfo('✅ ビューカード金額抽出テスト成功');
}

/**
 * ユニットテスト - 利用日抽出
 */
function testExtractUsageDate() {
    logInfo('=== 利用日抽出テスト開始 ===');

    const testCases = [
        {
            text: 'ご利用日時：2025/09/12 14:30',
            expected: '20250912',
            desc: 'ラベル付き（スラッシュ）'
        },
        {
            text: '利用日：2025-9-5',
            expected: '20250905',
            desc: 'ラベル付き（ハイフン）'
        },
        {
            text: 'ご利用日：2025年09月02日',
            expected: '20250902',
            desc: '和暦表記'
        },
        {
            text: '2025/12/25 クリスマス',
            expected: '20251225',
            desc: 'ラベルなし（フォールバック）'
        },
        {
            text: '日付なしテキスト',
            expected: null,
            desc: '日付なし'
        }
    ];

    testCases.forEach(tc => {
        const result = extractUsageDate(tc.text, '三井住友カード');
        const resultYmd = result ? toYmd(result) : null;
        assertEqual(resultYmd, tc.expected, `利用日抽出: ${tc.desc}`);
    });

    logInfo('✅ 利用日抽出テスト成功');
}

/**
 * ユニットテスト - 日付操作
 */
function testDateOperations() {
    logInfo('=== 日付操作テスト開始 ===');

    // addDays テスト
    const baseDate = new Date(2025, 0, 15);
    const plusThree = addDays(baseDate, 3);
    const minusTwo = addDays(baseDate, -2);

    assertEqual(toYmd(plusThree), '20250118', 'addDays: +3日');
    assertEqual(toYmd(minusTwo), '20250113', 'addDays: -2日');

    // inRangeInclusive テスト
    const start = new Date(2025, 0, 1);
    const end = new Date(2025, 0, 31);
    const dateInRange = new Date(2025, 0, 15);
    const dateOutRange = new Date(2025, 1, 1);

    assertEqual(inRangeInclusive(dateInRange, start, end), true, '範囲内判定: true');
    assertEqual(inRangeInclusive(dateOutRange, start, end), false, '範囲外判定: false');

    // compareYmd テスト
    const date1 = new Date(2025, 0, 10);
    const date2 = new Date(2025, 0, 20);
    const date3 = new Date(2025, 0, 10);

    assertEqual(compareYmd(date1, date2), -1, 'compareYmd: date1 < date2');
    assertEqual(compareYmd(date2, date1), 1, 'compareYmd: date2 > date1');
    assertEqual(compareYmd(date1, date3), 0, 'compareYmd: date1 == date3');

    logInfo('✅ 日付操作テスト成功');
}

/**
 * ユニットテスト - ビューカード件名分類
 */
function testClassifyViewSubject() {
    logInfo('=== ビューカード件名分類テスト開始 ===');

    const testCases = [
        {
            subject: '－確報版－ ビューカードご利用情報のお知らせ（本人会員利用）',
            expected: 'confirmed',
            desc: '確報版'
        },
        {
            subject: '－速報版－ ビューカードご利用情報のお知らせ',
            expected: 'provisional',
            desc: '速報版'
        },
        {
            subject: 'ビューカードご利用情報のお知らせ',
            expected: 'unknown',
            desc: 'バージョン指定なし'
        }
    ];

    testCases.forEach(tc => {
        const result = classifyViewSubject(tc.subject);
        assertEqual(result, tc.expected, `件名分類: ${tc.desc}`);
    });

    logInfo('✅ ビューカード件名分類テスト成功');
}

/**
 * ユニットテスト - テキスト正規化
 */
function testTextNormalization() {
    logInfo('=== テキスト正規化テスト開始 ===');

    // normalizeDigits テスト
    const fullWidth = '１２３４５';
    const normalized = normalizeDigits(fullWidth);
    assertEqual(normalized, '12345', '全角数字→半角数字');

    // sanitizeText テスト
    const messyText = '  テスト  \n\n\n改行  ';
    const sanitized = sanitizeText(messyText);
    assertEqual(sanitized, 'テスト \n改行', '空白・改行整理');

    logInfo('✅ テキスト正規化テスト成功');
}

/**
 * ユニットテスト - CONFIG検証
 */
function testConfigValidation() {
    logInfo('=== CONFIG検証テスト開始 ===');

    // CONFIG構造の存在確認
    assertNotNull(CONFIG, 'CONFIG存在確認');
    assertNotNull(CONFIG.GMAIL, 'CONFIG.GMAIL存在確認');
    assertNotNull(CONFIG.GMAIL.SOURCES, 'CONFIG.GMAIL.SOURCES存在確認');
    assertNotNull(CONFIG.GMAIL.SEARCH, 'CONFIG.GMAIL.SEARCH存在確認');
    assertNotNull(CONFIG.MESSAGE, 'CONFIG.MESSAGE存在確認');
    assertNotNull(CONFIG.CACHE, 'CONFIG.CACHE存在確認');

    // カード設定の検証
    assertEqual(CONFIG.GMAIL.SOURCES.length, 2, 'カード数: 2枚');

    const mitsui = CONFIG.GMAIL.SOURCES.find(s => s.NAME === '三井住友カード');
    assertNotNull(mitsui, '三井住友カード設定存在');
    assertEqual(mitsui.CYCLE_START, 1, '三井住友開始日: 1日');
    assertEqual(mitsui.CYCLE_END, 0, '三井住友終了日: 月末');

    const view = CONFIG.GMAIL.SOURCES.find(s => s.NAME === 'ビューカード');
    assertNotNull(view, 'ビューカード設定存在');
    assertEqual(view.CYCLE_START, 6, 'ビュー開始日: 6日');
    assertEqual(view.CYCLE_END, 5, 'ビュー終了日: 5日');

    // 検索設定の検証
    assertEqual(CONFIG.GMAIL.SEARCH.MAX_RESULTS, 500, '最大検索結果数');
    assertEqual(CONFIG.GMAIL.SEARCH.DATE_MARGIN_BEFORE, 2, '前マージン日数');
    assertEqual(CONFIG.GMAIL.SEARCH.DATE_MARGIN_AFTER, 2, '後マージン日数');

    logInfo('✅ CONFIG検証テスト成功');
}

/**
 * 全テスト実行
 */
function runAllTests() {
    logInfo('=============================');
    logInfo('   全テスト実行開始');
    logInfo('=============================');

    try {
        testConfigValidation();
        testCycleCalculation();
        testSumAmountsMitsui();
        testSumAmountsViewCard();
        testExtractUsageDate();
        testDateOperations();
        testClassifyViewSubject();
        testTextNormalization();

        logInfo('=============================');
        logInfo('   🎉 全テスト成功 🎉');
        logInfo('=============================');
    } catch (error) {
        logError('❌ テスト失敗:', error);
        throw error;
    }
}

/**
 * アサーションヘルパー: 値の等価性チェック
 * @param {*} actual - 実際の値
 * @param {*} expected - 期待値
 * @param {string} description - テストの説明
 */
function assertEqual(actual, expected, description) {
    if (actual !== expected) {
        const error = new Error(`${description}: 期待値=${expected}, 実際=${actual}`);
        logError('アサーション失敗', error);
        throw error;
    }
    logDebug(`✓ ${description}`);
}

/**
 * アサーションヘルパー: null/undefined でないことをチェック
 * @param {*} value - チェックする値
 * @param {string} description - テストの説明
 */
function assertNotNull(value, description) {
    if (value === null || value === undefined) {
        const error = new Error(`${description}: 値がnull/undefinedです`);
        logError('アサーション失敗', error);
        throw error;
    }
    logDebug(`✓ ${description}`);
}
