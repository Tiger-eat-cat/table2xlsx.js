import { Cell, Worksheet } from 'exceljs';
export declare const fontProcessor: (cell: HTMLTableCellElement, sheetCell: Cell, style: CSSStyleDeclaration) => void;
export declare const columnProcessor: (worksheet: Worksheet, from: number, to: number, cellStyle: CSSStyleDeclaration) => void;
export declare const hyperlinkProcessor: (cell: HTMLTableCellElement, sheetCell: Cell) => void;
export declare const inputProcessor: (cell: HTMLTableCellElement, sheetCell: Cell) => void;
export declare const imgProcessor: (cell: HTMLTableCellElement, sheetCell: Cell) => void;
