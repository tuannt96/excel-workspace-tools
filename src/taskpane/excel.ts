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

      // Tự động điều chỉnh độ rộng các cột
      range.format.autofitColumns();

      // Tự động điều chỉnh chiều cao các hàng
      range.format.autofitRows();

      await context.sync();

      console.log("Selected range has been automatically resized.");
    });
  } catch (error) {
    console.error(error);
  }
}