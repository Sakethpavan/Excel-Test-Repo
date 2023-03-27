import { Injectable } from '@angular/core';
import { Cell, Row, Workbook, Worksheet } from 'exceljs';
import * as fs from 'file-saver';
import {
  base64Img,
  convertToOperationSheetData,
  createCell,
  formatData,
} from './excel.helper';
import {
  ExcelProperties,
  OperationStepDetail,
  StandardOperationSheetData,
} from './excel.type';
import {
  cellStyles,
  centerAlignedCellStyles,
  rightAlignedCellStyles,
  tableHeaderCellStyles,
  tableValueCellStyles,
  titleStyles,
} from './styles-helper';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  constructor() {}

  download(data: any) {
    const operationSheetData: StandardOperationSheetData =
      convertToOperationSheetData(data);
    this.customExcel(operationSheetData, {
      worksheetName: 'Test',
      fileName: 'Sample',
    });
  }

  generateTitle(worksheet: Worksheet, title: string) {
    worksheet.mergeCells('A1:AA3');
    createCell({
      worksheet,
      cellLocation: 'A1',
      cellData: title,
      styles: titleStyles,
    });
  }

  async customExcel(
    operationSheetData: StandardOperationSheetData,
    excelProperties: ExcelProperties
  ) {
    const workbook: Workbook = new Workbook();
    const worksheet: Worksheet = workbook.addWorksheet(
      excelProperties.worksheetName,
      {
        pageSetup: {
          showGridLines: false,
          firstPageNumber: 1,
          orientation: 'landscape',
          paperSize: 9,
          fitToPage: true,
        },
        properties: {
          showGridLines: false,
        },
      }
    );

    // Row 1 to 3 title
    this.generateTitle(worksheet, 'Standard Operation Sheet - Procedure');

    // Logo
    const logo = workbook.addImage({
      base64: base64Img,
      extension: 'png',
    });

    worksheet.addImage(logo, {
      tl: { col: 23.5, row: 0.5 },
      br: { col: 26.5, row: 2.5 },
    });

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

    // Operation Number Cell
    createCell({
      worksheet,
      cellLocation: 'A4',
      cellData: 'Operation Number',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D4',
      cellData: operationSheetData.operationNumber,
      styles: cellStyles,
    });

    // Primary or Secondary Cell
    worksheet.getColumn(12).width = 21.36;
    createCell({
      worksheet,
      cellLocation: 'L4',
      cellData: 'Primary or Secondary',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'L5',
      cellData: operationSheetData.primarySecondarySos,
      styles: cellStyles,
    });

    // Prepared By Cell
    createCell({
      worksheet,
      cellLocation: 'R4',
      cellData: 'Prepared By',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T4',
      cellData: operationSheetData.preparedBy,
      styles: cellStyles,
    });

    // Applied Model Cell
    createCell({
      worksheet,
      cellLocation: 'V4',
      cellData: 'Applied Model',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X4',
      cellData: operationSheetData.appliedModel,
      styles: cellStyles,
    });

    // Revision Date
    createCell({
      worksheet,
      cellLocation: 'N5',
      cellData: 'Revision Date',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R5',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T5',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V5',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X5',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z5',
      cellData: '',
      styles: cellStyles,
    });
    /* Row 4 and Row 5  end*/

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

    // Operatrion Description
    createCell({
      worksheet,
      cellLocation: 'A6',
      cellData: 'Operation Description',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D6',
      cellData: operationSheetData.operationDescription,
      styles: cellStyles,
    });

    // Time to Master
    createCell({
      worksheet,
      cellLocation: 'K6',
      cellData: 'Time to Master',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'L6',
      cellData: operationSheetData.timeToMaster,
      styles: cellStyles,
    });

    // Issue Number
    createCell({
      worksheet,
      cellLocation: 'N6',
      cellData: 'Issue Number',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R6',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T6',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V6',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X6',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z6',
      cellData: '',
      styles: cellStyles,
    });

    // Revision Detail
    createCell({
      worksheet,
      cellLocation: 'N7',
      cellData: 'Revision Detail',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R7',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T7',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V7',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X7',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z7',
      cellData: '',
      styles: cellStyles,
    });
    /* Row 6 & 7 end*/

    /* Row 8 & 9 */
    worksheet.mergeCells('A8:C9');
    worksheet.mergeCells('D8:M9');
    worksheet.mergeCells('N8:N15');
    worksheet.mergeCells('O8:Q8');
    worksheet.mergeCells('O9:Q9');
    worksheet.mergeCells('R8:S8');
    worksheet.mergeCells('R9:S9');
    worksheet.mergeCells('T8:U8');
    worksheet.mergeCells('T9:U9');
    worksheet.mergeCells('V8:W8');
    worksheet.mergeCells('V9:W9');
    worksheet.mergeCells('X8:Y8');
    worksheet.mergeCells('X9:Y9');
    worksheet.mergeCells('Z8:AA8');
    worksheet.mergeCells('Z9:AA9');
    // PPE Requirements
    createCell({
      worksheet,
      cellLocation: 'A8',
      cellData: 'PPE Requirements',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D8',
      cellData: operationSheetData.ppeRequirements,
      styles: cellStyles,
    });

    // Senior Supervisor (1)
    createCell({
      worksheet,
      cellLocation: 'O8',
      cellData: 'Senior Supervisor (1)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R8',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T8',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V8',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X8',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z8',
      cellData: '',
      styles: cellStyles,
    });

    // Senior Supervisor (2)
    createCell({
      worksheet,
      cellLocation: 'O9',
      cellData: 'Senior Supervisor (2)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R9',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T9',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V9',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X9',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z9',
      cellData: '',
      styles: cellStyles,
    });
    /* Row 8 & 9 end*/

    /* Row 10 & 11 */
    worksheet.mergeCells('A10:C11');
    worksheet.mergeCells('D10:M11');
    worksheet.mergeCells('O10:P10');
    worksheet.mergeCells('O11:P11');
    worksheet.mergeCells('R10:S10');
    worksheet.mergeCells('R11:S11');
    worksheet.mergeCells('T10:U10');
    worksheet.mergeCells('T11:U11');
    worksheet.mergeCells('V10:W10');
    worksheet.mergeCells('V11:W11');
    worksheet.mergeCells('X10:Y10');
    worksheet.mergeCells('X11:Y11');
    worksheet.mergeCells('Z10:AA10');
    worksheet.mergeCells('Z11:AA11');
    // Jigs / Tools / Facility
    createCell({
      worksheet,
      cellLocation: 'A10',
      cellData: 'Jigs / Tools / Facility',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D10',
      cellData: '',
      styles: cellStyles,
    });

    // Supervisor (1)
    createCell({
      worksheet,
      cellLocation: 'O10',
      cellData: 'Supervisor',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Q10',
      cellData: '(1)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R10',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T10',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V10',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X10',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z10',
      cellData: '',
      styles: cellStyles,
    });

    // Supervisor (2)
    createCell({
      worksheet,
      cellLocation: 'O11',
      cellData: 'Supervisor',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Q11',
      cellData: '(2)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R11',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T11',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V11',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X11',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z11',
      cellData: '',
      styles: cellStyles,
    });

    /* Row 10 & 11 end*/

    /* Row 12 & 13 */
    worksheet.mergeCells('A12:C13');
    worksheet.mergeCells('D12:M13');
    worksheet.mergeCells('O12:P12');
    worksheet.mergeCells('O13:P13');
    worksheet.mergeCells('R12:S12');
    worksheet.mergeCells('R13:S13');
    worksheet.mergeCells('T12:U12');
    worksheet.mergeCells('T13:U13');
    worksheet.mergeCells('V12:W12');
    worksheet.mergeCells('V13:W13');
    worksheet.mergeCells('X12:Y12');
    worksheet.mergeCells('X13:Y13');
    worksheet.mergeCells('Z12:AA12');
    worksheet.mergeCells('Z13:AA13');
    // Significant Hazards
    createCell({
      worksheet,
      cellLocation: 'A12',
      cellData: 'Significant Hazards',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D12',
      cellData: operationSheetData.significantHazard,
      styles: cellStyles,
    });

    // Supervisor (3)
    createCell({
      worksheet,
      cellLocation: 'O12',
      cellData: 'Supervisor',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Q12',
      cellData: '(3)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R12',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T12',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V12',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X12',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z12',
      cellData: '',
      styles: cellStyles,
    });

    // Supervisor (4)
    createCell({
      worksheet,
      cellLocation: 'O13',
      cellData: 'Supervisor',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Q13',
      cellData: '(4)',
      styles: centerAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'R13',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'T13',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'V13',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'X13',
      cellData: '',
      styles: cellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'Z13',
      cellData: '',
      styles: cellStyles,
    });

    /* Row 12 & 13 end*/

    /* Row 14 & 15 */
    worksheet.mergeCells('A14:C15');
    worksheet.mergeCells('D14:M15');
    worksheet.mergeCells('O14:AA15');
    // Materials Used
    createCell({
      worksheet,
      cellLocation: 'A14',
      cellData: 'Materials Used',
      styles: rightAlignedCellStyles,
    });

    createCell({
      worksheet,
      cellLocation: 'D14',
      cellData: operationSheetData.materialsUsed,
      styles: cellStyles,
    });

    // Revision signatory
    createCell({
      worksheet,
      cellLocation: 'N8',
      cellData: 'Revision signatory',
      styles: {
        ...cellStyles,
        alignment: {
          horizontal: 'center',
          vertical: 'middle',
          wrapText: true,
          textRotation: 90,
        },
        border: {
          ...cellStyles.border,
          right: {
            style: 'thin',
            color: {
              argb: 'FFFFFFFF',
            },
          },
        },
      },
    });

    createCell({
      worksheet,
      cellLocation: 'O14',
      cellData: '',
      styles: {
        ...cellStyles,
        border: {
          ...cellStyles.border,
          left: {
            style: 'thin',
            color: {
              argb: '00FFFFFF',
            },
          },
        },
      },
    });
    /* Row 14 & 15 end*/

    /* Row 16 */
    this.generateTable(worksheet, operationSheetData.operationStepDetails);
    /* Row 16 end*/
    this.downloadFile(workbook, excelProperties.fileName);
  }

  generateTable(worksheet: Worksheet, data: OperationStepDetail[]) {
    worksheet.mergeCells('B16:F16');
    worksheet.mergeCells('H16:L16');
    worksheet.mergeCells('M16:AA16');
    // No
    createCell({
      worksheet,
      cellLocation: 'A16',
      cellData: 'No',
      styles: tableHeaderCellStyles,
    });

    // Main Steps
    createCell({
      worksheet,
      cellLocation: 'B16',
      cellData: 'Main Steps',
      styles: tableHeaderCellStyles,
    });

    // Time
    createCell({
      worksheet,
      cellLocation: 'G16',
      cellData: 'Time',
      styles: tableHeaderCellStyles,
    });

    // Key Points and (Reasons)
    createCell({
      worksheet,
      cellLocation: 'H16',
      cellData: 'Key Points and (Reasons)',
      styles: tableHeaderCellStyles,
    });

    // Operation Routes \ Illustrations \ Specifications
    createCell({
      worksheet,
      cellLocation: 'M16',
      cellData: 'Operation Routes \\ Illustrations \\ Specifications',
      styles: tableHeaderCellStyles,
    });

    data.forEach((rowData: OperationStepDetail, index: number) => {
      const newRowValues = [];
      newRowValues['A'.charCodeAt(0) - '@'.charCodeAt(0)] = index + 1;
      newRowValues['B'.charCodeAt(0) - '@'.charCodeAt(0)] = formatData(
        rowData.stepDescription
      );
      newRowValues['G'.charCodeAt(0) - '@'.charCodeAt(0)] = formatData(
        rowData.stepTime
      );
      newRowValues['H'.charCodeAt(0) - '@'.charCodeAt(0)] = formatData(
        rowData.keyPoint
      );
      newRowValues['M'.charCodeAt(0) - '@'.charCodeAt(0)] = formatData(
        rowData.operationAnalysis
      );
      const row: Row = worksheet.addRow(newRowValues);
      worksheet.mergeCells(`B${row.number}:F${row.number}`);
      worksheet.mergeCells(`H${row.number}:L${row.number}`);
      worksheet.mergeCells(`M${row.number}:AA${row.number}`);

      row.eachCell((cell: Cell) => {
        cell.style = tableValueCellStyles;
      });
    });
  }

  async downloadFile(workbook: Workbook, outputFileName: string) {
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      fs.saveAs(blob, outputFileName);
    });
  }
}
