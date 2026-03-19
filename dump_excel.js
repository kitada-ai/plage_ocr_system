const XlsxPopulate = require("xlsx-populate");
const path = require("path");

async function dump() {
    const templatePath = path.join(process.cwd(), "docs", "samples", "sheets", "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx");
    const workbook = await XlsxPopulate.fromFileAsync(templatePath);
    const sheet = workbook.sheet("申込書");
    console.log("--- Template Dump (Sheet: 申込書) ---");
    for (let r = 1; r <= 25; r++) {
        const row = [];
        for (let c = 1; c <= 20; c++) {
            row.push(sheet.row(r).cell(c).value() || "");
        }
        console.log(`Row ${r}: ${row.join(" | ")}`);
    }
}

dump().catch(console.error);
