import { Cell } from 'exceljs'
import { rgbToArgb } from './tools'
import { TEXT_ALIGN } from './config'

export const fontProcessor = (cell: HTMLTableCellElement, sheetCell: Cell): void => {
    const style: CSSStyleDeclaration = getComputedStyle(cell)
    const fontSize: string = style.fontSize
    const textAlign: string = style.textAlign
    const color: string = rgbToArgb(style.color)
    sheetCell.font = {
        size: parseInt(fontSize),
        color: { argb: color },
    }
    if (TEXT_ALIGN.some(item => item === textAlign)) {
        sheetCell.alignment = {
            // @ts-ignore
            horizontal: textAlign,
        }
    }
}
