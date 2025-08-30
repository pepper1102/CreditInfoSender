
/**
 * 金額抽出ロジックのテスト
 */
function test_sumAmountsAfterMarker() {
  const sampleText = `
    ご利用内容
    ○○ショップ 1,200円
    △△ストア 3,400円
    テスト 100円
  `;
  const result = sumAmountsMitsui(sampleText);
  Logger.log("期待値=4700, 実際=%s", result);
}

/**
 * メッセージ生成ロジックのテスト
 */
function test_messageFormat() {
  const today = "2025/08/22";
  const monthTotal = 15000;
  const diff = -3000; // 前日より減った場合を想定
  const diffSign = diff >= 0 ? '+' : '-';
  const diffAbs = Math.abs(diff);

  const message =
    `${today}時点でのカード利用額は\n` +
    `¥${formatJPY(monthTotal)}（前日より${diffSign}¥${formatJPY(diffAbs)}）です。\n\n` +
    `常に詳細を表示する`;

  Logger.log(message);
}
