import { Buffer, Workbook } from 'exceljs'

export interface Excel {
    export(filename: string): Promise<Buffer>
    workbook?: Workbook
}
