const ExcelJS = require("exceljs");

exports.handler = async () => {
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Matrix Systems";
    workbook.created = new Date();

    const ws = workbook.addWorksheet("Template", {
      views: [{ state: "frozen", ySplit: 1 }], // freeze header row
    });

    // Columns from your screenshot
    const columns = [
      { header: "ID #", key: "id", width: 10 },
      { header: "LINE 1 TEXT", key: "line1", width: 24 },
      { header: "LINE 2 TEXT", key: "line2", width: 24 },
      { header: "LINE 3 TEXT", key: "line3", width: 24 },
      { header: "LINE 1 LETTER HEIGHT", key: "l1h", width: 22 },
      { header: "LINE 2 LETTER HEIGHT", key: "l2h", width: 22 },
      { header: "LINE 3 LETTER HEIGHT", key: "l3h", width: 22 },
      { header: "QTY", key: "qty", width: 10 },
      { header: "BACKGRND COLOR", key: "bg", width: 18 },
      { header: "LETTER COLOR", key: "lc", width: 16 },
      { header: "WIDTH (INCHES)", key: "w", width: 16 },
      { header: "HEIGHT (INCHES)", key: "h", width: 16 },
      { header: "STICKY BACK", key: "sticky", width: 14 },
      { header: "COMMENTS", key: "comments", width: 28 },
    ];

    ws.columns = columns;

    // Header styling
    const headerRow = ws.getRow(1);
    headerRow.height = 20;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2B2B2B" } };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    });

    // Optional: add a few blank rows so it “looks like a form”
    for (let i = 0; i < 25; i++) ws.addRow({});

    // Generate file
    const buffer = await workbook.xlsx.writeBuffer();
    const b64 = Buffer.from(buffer).toString("base64");

    return {
      statusCode: 200,
      isBase64Encoded: true,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="matrix_template.xlsx"',
        "Cache-Control": "no-store",
      },
      body: b64,
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Failed to generate Excel file", details: String(err) }),
    };
  }
};
