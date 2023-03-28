import { Cell } from 'exceljs';
import { CellProperties, StandardOperationSheetData } from './excel.type';

// creates a cell at given cell location with applied styles
export const createCell = ({
  worksheet,
  cellLocation,
  cellData,
  styles,
}: CellProperties): Cell => {
  const cell: Cell = worksheet.getCell(cellLocation);
  cell.value = cellData;
  if (styles) {
    Object.entries(styles).forEach(([styleName, styleValue]) => {
      (cell as any)[styleName] = styleValue;
    });
  }
  return cell;
};

export const convertToOperationSheetData = (
  data: any
): StandardOperationSheetData => {
  const operationData: StandardOperationSheetData = {
    operationId: formatData<string>(data.operationId),
    operationNumber: formatData<string>(data.operationNumber),
    primarySecondarySos: formatData<string>(data.primarySecondarySos),
    operationDescription: formatData<string>(data.operationDescription),
    timeToMaster: formatData<string>(''),
    tools: formatData<string>(data.toolsRequired),
    ppeRequirements: formatData<string>(data.ppeRequirements),
    significantHazard: formatData<string>(data.significantHazard),
    materialsUsed: formatData<string>(data.materialsRequired), // mapping to materials Required
    operationStepDetails: data.operationStepDetails, // OperationStep[]
    preparedBy: formatData<string>(data.preparedBy),
    appliedModel: formatData<string>(data.appliedModel),
    total: formatData<string>(data.duration),
  };

  return operationData;
};

// checks if data is avaiable or not
export function formatData<T>(value: T | undefined | null): T | string {
  if (value == null || value == undefined) return '';
  return value;
}

export const base64Img = '';
