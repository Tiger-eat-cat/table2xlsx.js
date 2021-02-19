import { Workbook } from 'exceljs'

export interface Excel {
    export(filename: string): Promise<unknown>
    workbook?: Workbook
}
