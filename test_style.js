const XlsxPopulate = require("xlsx-populate");

async function test() {
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0);
    try {
        sheet.cell("A1").style("fontFamily", "Yu Gothic");
        console.log("fontFamily worked");
    } catch (e) {
        console.log("fontFamily failed:", e.message);
    }
    
    try {
        sheet.cell("A1").style("fontName", "Yu Gothic");
        console.log("fontName worked");
    } catch (e) {
        console.log("fontName failed:", e.message);
    }
}

test();
