const XlsxPopulate = require("xlsx-populate");
const path = require("path");

async function dump() {
    const templatePath = path.join(process.cwd(), "docs", "samples", "written", "テンプレ.xlsx");
    const workbook = await XlsxPopulate.fromFileAsync(templatePath);
    console.log("Sheets:", workbook.sheets().map(s => s.name()));
    const sheet = workbook.sheet(0); // Use index 0 as fallback
    console.log(`--- '${sheet.name()}' Dump ---`);
    for (let r = 1; r <= 30; r++) {
        const row = [];
        for (let c = 1; c <= 22; c++) {
            row.push(sheet.row(r).cell(c).value() || "");
        }
        console.log(`Row ${r}: ${row.join(" | ")}`);
    }
}

dump().catch(console.error);
