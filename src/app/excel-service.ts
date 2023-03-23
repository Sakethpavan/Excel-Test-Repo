import { Injectable } from '@angular/core';
import {
  Alignment,
  Cell,
  Font,
  Style,
  Workbook,
  Worksheet,
} from 'exceljs';
import * as fs from 'file-saver';
type CellProperties = {
  worksheet: Worksheet;
  cellLocation: string;
  cellData: string;
  styles?: Partial<Style>;
};
@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  constructor() {}

  download() {
    this.customExcel();
  }

  createCell({
    worksheet,
    cellLocation,
    cellData,
    styles,
  }: CellProperties): Cell {
    const cell: Cell = worksheet.getCell(cellLocation);
    cell.value = cellData;
    if (styles) {
      Object.entries(styles).forEach(([styleName, styleValue]) => {
        (cell as any)[styleName] = styleValue;
      });
    }
    return cell;
  }
  async customExcel() {
    const workbook: Workbook = new Workbook();
    const worksheet: Worksheet = workbook.addWorksheet('Test');

    // Add title Row
    const titleFont: Partial<Font> = {
      name: 'Arial',
      size: 26,
      bold: true,
      underline: true,
    };
    const titleAlignment: Partial<Alignment> = {
      horizontal: 'center',
      vertical: 'middle',
    };
    const titleStyles: Partial<Style> = {
      font: titleFont,
      alignment: titleAlignment,
    };

    worksheet.mergeCells('A1:AA3');
    this.createCell({
      worksheet,
      cellLocation: 'A1',
      cellData: 'Standard Operation Sheet - Procedure',
      styles: titleStyles,
    });

    // operation Number row
    const cellStyles: Partial<Style> = {
      font: {
        name: 'Arial',
        size: 10,
        bold: true,
      },
      alignment: {
        horizontal: 'right',
        vertical: 'middle',
      },
    };

   

    /* Row 4 and Row 5 */
    worksheet.mergeCells('A4:C5');
    worksheet.mergeCells('D4:K5');
    worksheet.mergeCells('L4:M4');
    worksheet.mergeCells('L5:M5');
    worksheet.mergeCells('N4:Q4');
    worksheet.mergeCells('N5:Q5');
    worksheet.mergeCells('R4:S4');
    worksheet.mergeCells('R5:S5');
    worksheet.mergeCells('T4:U4');
    worksheet.mergeCells('T5:U5');
    worksheet.mergeCells('V4:W4');
    worksheet.mergeCells('V5:W5');
    worksheet.mergeCells('X4:AA4');
    worksheet.mergeCells('X5:Y5');
    worksheet.mergeCells('Z5:AA5');

    // Operator Number Cell
    this.createCell({
      worksheet,
      cellLocation: 'A4',
      cellData: 'Operator Number',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'D4',
      cellData: '1000',
      styles: cellStyles,
    });

    // Primary or Secondary Cell
    worksheet.getColumn(12).width = 21.36;
    this.createCell({
      worksheet,
      cellLocation: 'L4',
      cellData: 'Primary or Secondary',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'L5',
      cellData: '',
      styles: cellStyles,
    });

    // Prepared By Cell
    this.createCell({
      worksheet,
      cellLocation: 'R4',
      cellData: 'Prepared By',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'T4',
      cellData: 'Saketh pavan',
      styles: cellStyles,
    });

    // Applied Model Cell
    this.createCell({
      worksheet,
      cellLocation: 'V4',
      cellData: 'Applied Model',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'X4',
      cellData: '**Applied Model Value ***',
      styles: cellStyles,
    });

    // Revision Date
    this.createCell({
      worksheet,
      cellLocation: 'N5',
      cellData: 'Revision Date',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'R5',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'T5',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'V5',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'X5',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'Z5',
      cellData: '',
      styles: cellStyles,
    });


    /* Row 6 & 7 */
    worksheet.mergeCells('A6:C7');
    worksheet.mergeCells('D6:J7');
    worksheet.getColumn(11).width = 13.91;
    worksheet.mergeCells('K6:K7');
    worksheet.mergeCells('L6:M7');
    worksheet.mergeCells('N6:Q6');
    worksheet.mergeCells('N7:Q7');
    worksheet.mergeCells('R6:S6');
    worksheet.mergeCells('R7:S7');
    worksheet.mergeCells('T6:U6');
    worksheet.mergeCells('T7:U7');
    worksheet.mergeCells('V6:W6');
    worksheet.mergeCells('V7:W7');
    worksheet.mergeCells('X6:Y6');
    worksheet.mergeCells('X7:Y7');
    worksheet.mergeCells('Z6:AA6');
    worksheet.mergeCells('Z7:AA7');

    // Operator Description
    this.createCell({
      worksheet,
      cellLocation: 'A6',
      cellData: 'Operator Description',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'D6',
      cellData: 'Some operator description',
      styles: cellStyles,
    });

    // Time to Master
    this.createCell({
      worksheet,
      cellLocation: 'K6',
      cellData: 'Time to Master',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'L6',
      cellData: '',
      styles: cellStyles,
    });

    // Time to Master
    this.createCell({
      worksheet,
      cellLocation: 'K6',
      cellData: 'Time to Master',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'L6',
      cellData: 'Sometime to master',
      styles: cellStyles,
    });

    // Issue Number
    this.createCell({
      worksheet,
      cellLocation: 'N6',
      cellData: 'Issue Number',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'R6',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'T6',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'V6',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'X6',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'Z6',
      cellData: '',
      styles: cellStyles,
    });

    // Revision Detail
    this.createCell({
      worksheet,
      cellLocation: 'N7',
      cellData: 'Revision Detail',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'R7',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'T7',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'V7',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'X7',
      cellData: '',
      styles: cellStyles,
    });

    this.createCell({
      worksheet,
      cellLocation: 'Z7',
      cellData: '',
      styles: cellStyles,
    });

    

    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      fs.saveAs(blob, 'Client.xlsx');
    });
  }
}
