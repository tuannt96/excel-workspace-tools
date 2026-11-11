/*
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.querySelector("#run").onclick = runExcel;
  }
});

export async function runExcel() {
  try {
    await Excel.run(async (context) => {
      // Lấy vùng đang được chọn
      const range = context.workbook.getSelectedRange();

      // Lấy giá trị của các ô
      range.load("values");

      await context.sync();

      // Chuyển các giá trị dạng chuỗi thành chữ in hoa
      const newValues = range.values.map((row) =>
        row.map((value) => (typeof value === "string" ? value.toUpperCase() : value))
      );

      // Ghi lại dữ liệu đã chuyển đổi
      range.values = newValues;

      await context.sync();

      console.log("Selected range converted to uppercase.");
    });
  } catch (error) {
    console.error(error);
  }
}
