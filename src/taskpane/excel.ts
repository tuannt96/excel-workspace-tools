/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runExcel;
  }
});

export async function runExcel() {
  try {
    await Excel.run(async (context) => {
      // Lấy vùng đang được chọn
      const range = context.workbook.getSelectedRange();

      // Đọc dữ liệu trong vùng chọn
      range.load(["values", "rowCount", "columnCount"]);

      await context.sync();

      const values = range.values;
      let emptyCellCount = 0;

      // Duyệt từng ô trong vùng chọn
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const cell = range.getCell(r, c);
          const value = values[r][c];

          // Nếu ô rỗng hoặc chỉ chứa khoảng trắng
          if (
            value === null ||
            value === "" ||
            (typeof value === "string" && value.trim() === "")
          ) {
            cell.format.fill.color = "#FDE7E9";
            emptyCellCount++;
          }
        }
      }

      await context.sync();

      console.log(`Highlighted ${emptyCellCount} empty cell(s).`);
    });
  } catch (error) {
    console.error(error);
  }
}