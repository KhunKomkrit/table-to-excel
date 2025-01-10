import ExcelJS from "exceljs";
import { JSDOM } from "jsdom";

interface MergedCellData {
  value: string;
  style: CSSStyleDeclaration;
}

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
      
      // First pass: Create all rows and cells with basic values
      let rowNumber = worksheet.rowCount + 1;
      const mergeCells: Array<{
        startRow: number;
        startCol: number;
        endRow: number;
        endCol: number;
        value: string;
        style: CSSStyleDeclaration;
      }> = [];

      const occupiedCells = new Set<string>();

      rows.forEach((row, rowIndex) => {
        const newRow = worksheet.addRow([]);
        const cells = row.querySelectorAll('th, td');
        let colIndex = 1;

        cells.forEach((cell) => {

          while (occupiedCells.has(`${rowNumber + rowIndex},${colIndex}`)) {
            colIndex++;
          }

          const value = cell.textContent?.trim() || '';
          const style = (cell as HTMLElement).style;
          const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
          const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);

          const excelCell = newRow.getCell(colIndex);
          excelCell.value = value;
          applyStyles(excelCell, style);

          if (colspan > 1 || rowspan > 1) {
            mergeCells.push({
              startRow: rowNumber + rowIndex,
              startCol: colIndex,
              endRow: rowNumber + rowIndex + rowspan - 1,
              endCol: colIndex + colspan - 1,
              value,
              style
            });

            for (let r = 0; r < rowspan; r++) {
              for (let c = 0; c < colspan; c++) {
                occupiedCells.add(`${rowNumber + rowIndex + r},${colIndex + c}`);
              }
            }
          }

          colIndex += colspan;
        });
      });

      mergeCells.forEach(({ startRow, startCol, endRow, endCol, value, style }) => {
        try {
          worksheet.mergeCells(startRow, startCol, endRow, endCol);
          const cell = worksheet.getCell(startRow, startCol);
          cell.value = value;
          applyStyles(cell, style);
        } catch (error) {
          console.warn(`Failed to merge cells from (${startRow}, ${startCol}) to (${endRow}, ${endCol}):`, error);
        }
      });

      if (tableIndex < tables.length - 1) {
        worksheet.addRow([]);
      }
    });
  
    return workbook;
  }

  function applyStyles(excelCell: ExcelJS.Cell, style: CSSStyleDeclaration) {
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
}
