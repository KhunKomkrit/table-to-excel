import ExcelJS from "exceljs";
import { JSDOM } from "jsdom";

export async function htmlToExcel(
  html: string,
  outputFilePath: string
): Promise<void> {
  const workbook = generateHtmlTable(html);
  await workbook.xlsx.writeFile(outputFilePath);
}

export async function htmlToExcelBlob(html: string): Promise<Blob> {
  const workbook = generateHtmlTable(html);
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function generateHtmlTable(html: string): ExcelJS.Workbook {
  const dom = new JSDOM(html);
  const table = dom.window.document.querySelector("table");
  if (!table) {
    throw new Error("No table element found in the HTML.");
  }

  const rows = table.querySelectorAll("tr");
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Styled Sheet");

  rows.forEach((row) => {
    const cells = row.querySelectorAll("th, td");
    const newRow = worksheet.addRow([]);

    cells.forEach((cell, colIndex) => {
      const value = cell.textContent?.trim() || "";
      const style = (cell as HTMLElement).style;

      const excelCell = newRow.getCell(colIndex + 1);
      excelCell.value = value;

      if (style.backgroundColor) {
        const color = style.backgroundColor
          .replace("rgb(", "")
          .replace(")", "")
          .split(",");
        excelCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb: `FF${parseInt(color[0])
              .toString(16)
              .padStart(2, "0")}${parseInt(color[1])
              .toString(16)
              .padStart(2, "0")}${parseInt(color[2])
              .toString(16)
              .padStart(2, "0")}`,
          },
        };
      }

      if (style.color) {
        const color = style.color
          .replace("rgb(", "")
          .replace(")", "")
          .split(",");
        excelCell.font = {
          color: {
            argb: `FF${parseInt(color[0])
              .toString(16)
              .padStart(2, "0")}${parseInt(color[1])
              .toString(16)
              .padStart(2, "0")}${parseInt(color[2])
              .toString(16)
              .padStart(2, "0")}`,
          },
        };
      }
    });
  });
  return workbook;
}
