import { Workbook } from 'exceljs';
export interface Excel {
    export(filename: string): Promise<unknown>;
    workbook?: Workbook;
}
export declare enum TagName {
    img = "IMG",
    input = "INPUT",
    hyperlink = "A"
}
