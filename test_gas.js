// ローカル検証用スクリプト
function testGanttPositions() {
    const headers = ["ステータス", "セクション", "項目", "内容", "担当者", "決定/未定", "開始日", "終了日", "所要日数", "備考", "TaskID", "セパレータ"];
    console.log("=== ヘッダー ===");
    headers.forEach((h, i) => console.log(`${i + 1}列目 (${String.fromCharCode(65 + i)}列): ${h}`));

    console.log("\nカレンダー開始列 (headers.length + 1) =", headers.length + 1, "->", String.fromCharCode(65 + headers.length) + "列");

    console.log("\n=== 想定している書き込み位置 ===");
    console.log("H列(終了日) -> setFormula", "I2"); // <= これがおかしい。Hは8列目。数式はI列(9列目)?
}
testGanttPositions();
