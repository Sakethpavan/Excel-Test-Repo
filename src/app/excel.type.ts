import { Style, Worksheet } from 'exceljs';

export type CellProperties = {
  worksheet: Worksheet;
  cellLocation: string;
  cellData: string;
  styles?: Partial<Style>;
};

export type StandardOperationSheetData = {
  operationId: string;
  operationNumber: string;
  primarySecondarySos: string;
  operationDescription: string;
  timeToMaster: string; // mapping of duration
  ppeRequirements: string;
  significantHazard: string;
  materialsUsed: string; // mapping to materials Required
  operationStepDetails: OperationStep[];
  preparedBy: string;
  appliedModel: string;
};

export type OperationStep = {
  stepid: string;
  stepSequence: number;
  stepDescription: string;
  stepTime: string;
  stepShapeType: string;
  keyPoint: string;
  operationAnalysis: string | null;
  specialCharacteristicsDetail: {
    specialCharacteristicsId: string;
    specialCharacteristics: string;
  };
};


export type ExcelProperties = {
  worksheetName: string;
  fileName: string;
}