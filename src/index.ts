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
    const tables = dom.window.document.querySelectorAll('table');
  
    if (!tables.length) {
      throw new Error('No table elements found in the HTML.');
    }
  
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Worksheet');
  
    tables.forEach((table, tableIndex) => {
      const rows = table.querySelectorAll('tr');
  
      rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('th, td');
        const newRow = worksheet.addRow([]);
  
        cells.forEach((cell, colIndex) => {
          const value = cell.textContent?.trim() || '';
          const style = (cell as HTMLElement).style;

          const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
          const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
  
          const excelCell = newRow.getCell(colIndex + 1);
          excelCell.value = value;
  
          if (style.backgroundColor) {
            const color = style.backgroundColor.replace('rgb(', '').replace(')', '').split(',');
            excelCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: {
                argb: `FF${parseInt(color[0]).toString(16).padStart(2, '0')}${parseInt(color[1])
                  .toString(16)
                  .padStart(2, '0')}${parseInt(color[2]).toString(16).padStart(2, '0')}`,
              },
            };
          }
  
          if (style.color) {
            const color = style.color.replace('rgb(', '').replace(')', '').split(',');
            excelCell.font = {
              color: {
                argb: `FF${parseInt(color[0]).toString(16).padStart(2, '0')}${parseInt(color[1])
                  .toString(16)
                  .padStart(2, '0')}${parseInt(color[2]).toString(16).padStart(2, '0')}`,
              },
            };
          }

          if (style.border || style.borderWidth) {
            excelCell.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'thin', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } },
            };
          }

          if (colspan > 1 || rowspan > 1) {
            const startRow = newRow.number;
            const startCol = colIndex + 1;
            const endRow = startRow + rowspan - 1;
            const endCol = startCol + colspan - 1;
      
            worksheet.mergeCells(startRow, startCol, endRow, endCol);
          }

          if (style.textAlign) {
            const alignmentMap: Record<string, 'left' | 'center' | 'right'> = {
              left: 'left',
              center: 'center',
              right: 'right',
            };
        
            const alignValue = alignmentMap[style.textAlign as keyof typeof alignmentMap];
            if (alignValue) {
              excelCell.alignment = { horizontal: alignValue };
            }
          }

        });
  
        if (rowIndex === rows.length - 1 && tableIndex < tables.length - 1) {
          worksheet.addRow([]);
        }
      });
    });
  
    return workbook;
  }
