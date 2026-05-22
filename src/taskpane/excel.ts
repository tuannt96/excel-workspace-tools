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
      const range = context.workbook.getSelectedRange();

      range.format.fill.color = "#EAF4FF";
      range.format.font.color = "#1F2937";
      range.format.font.name = "Aptos";
      range.format.font.size = 11;

      range.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
      range.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
      range.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
      range.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
      range.format.borders.getItem("InsideHorizontal").style = Excel.BorderLineStyle.continuous;
      range.format.borders.getItem("InsideVertical").style = Excel.BorderLineStyle.continuous;

      range.format.autofitColumns();
      range.format.autofitRows();

      await context.sync();
      console.log("Selected range formatted successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}
