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
    const CONVERT_RATIO = 0.35
    if (from === to && cellStyle.width !== 'auto') {
        worksheet.getColumn(from).width = parseFloat(cellStyle.width) * CONVERT_RATIO
    }
}

export const hyperlinkProcessor = (cell: HTMLTableCellElement, sheetCell: Cell) => {
    const children = cell.children
    const tagName = children[0]?.tagName.toUpperCase()
    if (tagName === 'A') {
        const hyperlink = children[0] as HTMLLinkElement
        sheetCell.value = { text: cell.innerText, hyperlink: hyperlink.href }
    }
}

export const inputProcessor = (cell: HTMLTableCellElement, sheetCell: Cell) => {
    const children = cell.children
    const tagName = children[0]?.tagName.toUpperCase()
    if (tagName === 'INPUT') {
        const input = children[0] as HTMLInputElement
        sheetCell.value = input.value
    }
}
