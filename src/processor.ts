import { Cell, Worksheet } from 'exceljs'
import { rgbToArgb } from './tools'
import { TEXT_ALIGN } from './config'

export const fontProcessor = (cell: HTMLTableCellElement, sheetCell: Cell, style: CSSStyleDeclaration): void => {
    const fontSize: string = style.fontSize
    const textAlign: string = style.textAlign
    const color: string = rgbToArgb(style.color)
    const fontWeight = style.fontWeight
    const BOLD = 700
    sheetCell.font = {
        size: parseInt(fontSize),
        color: { argb: color },
        italic: style.fontStyle === 'italic',
        bold: fontWeight === 'bold' || parseInt(fontWeight) >= BOLD
    }
    if (TEXT_ALIGN.some(item => item === textAlign)) {
        sheetCell.alignment = {
            // @ts-ignore
            horizontal: textAlign,
        }
    }
}

export const columnProcessor = (worksheet: Worksheet, from: number, to: number, cellStyle: CSSStyleDeclaration): void => {
    if (from === to && cellStyle.width !== 'auto') {
        worksheet.getColumn(from).width = parseFloat(cellStyle.width) * 0.14
        console.log('=======================>', worksheet.getColumn(from))
    }
}
