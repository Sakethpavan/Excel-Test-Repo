import { Alignment, Border, Cell, Font, Style } from 'exceljs';

export const titleFont: Partial<Font> = {
  name: 'Arial',
  size: 26,
  bold: true,
  underline: true,
};

export const titleAlignment: Partial<Alignment> = {
  horizontal: 'center',
  vertical: 'middle',
};

export const titleStyles: Partial<Style> = {
  font: titleFont,
  alignment: titleAlignment,
};

export const thinBlackBorderStyle: Partial<Border> = {
  style: 'thin',
  color: {
    argb: 'FF000000',
  },
};

export const cellStyles: Partial<Style> = {
  font: {
    name: 'Arial',
    size: 10,
    bold: true,
  },
  alignment: {
    horizontal: 'left',
    vertical: 'middle',
  },
  border: {
    top: thinBlackBorderStyle,
    left: thinBlackBorderStyle,
    bottom: thinBlackBorderStyle,
    right: thinBlackBorderStyle,
  },
};

export const rightAlignedCellStyles: Partial<Style> = {
  ...cellStyles,
  alignment: {
    horizontal: 'right',
    vertical: 'middle',
  },
};

export const centerAlignedCellStyles: Partial<Style> = {
  ...cellStyles,
  alignment: {
    horizontal: 'center',
    vertical: 'middle',
  },
};

export const tableHeaderCellStyles: Partial<Style> = {
  ...cellStyles,
  font: {
    name: 'Arial',
    size: 14,
    bold: true,
  },
  alignment: {
    horizontal: 'center',
    vertical: 'middle',
  },
};

