import ExcelJS, { Cell } from 'exceljs'
import { fontProcessor, columnProcessor, hyperlinkProcessor, inputProcessor } from './processor'
import { Excel } from './types'

export const createExcel = (selector: string | HTMLTableElement [] = 'table'): Excel => {
    const workbook = new ExcelJS.Workbook()
    const tableElements  = typeof selector === 'string' ? document.querySelectorAll(selector) : selector
    const tables = Array.from(tableElements) as HTMLTableElement[]
    tables.forEach((table, index) => {
        const worksheet = workbook.addWorksheet(`Worksheet${index + 1}`)
        const rows: HTMLTableRowElement [] = Array.from(table.rows)
        const sheetHeight = rows.length
        const sheetWidth = rows[sheetHeight - 1].cells.length
        const generateRow = (): boolean [] => new Array(sheetWidth).fill(false)
        const mergeLog: boolean[][] = new Array(sheetHeight).fill(null).map(() => generateRow())
        rows.forEach((row, rowIndex) => {
            const y = rowIndex + 1 // 纵坐标
            let x = 1 // 横坐标
            const currentLineLog = mergeLog[rowIndex]
            for (let i = 0; i < sheetWidth; i++) {
                if (!currentLineLog[i]) {
                    x = i + 1
                    break
                }
            }
            const cells: HTMLTableCellElement [] = Array.from(row.cells)
            cells.forEach(cell => {
                const { colSpan, rowSpan } = cell
                const top = y // 开始行
                const left = x // 开始列
                const bottom = y + rowSpan - 1 // 结束行
                const right = x + colSpan - 1 // 结束列
                worksheet.mergeCells(top, left, bottom, right)
                const sheetCell: Cell = worksheet.getCell(top, left)
                sheetCell.value = cell.innerText
                const style: CSSStyleDeclaration = getComputedStyle(cell)
                fontProcessor(cell, sheetCell, style)
                columnProcessor(worksheet, left, right, style)
                hyperlinkProcessor(cell, sheetCell)
                inputProcessor(cell, sheetCell)
                for (let i = top - 1; i < bottom; i ++) {
                    for (let j = left - 1; j < right; j++) {
                        mergeLog[i][j] = true
                    }
                }
                x += colSpan
            })
        })
    })
    return {
        export: async (filename: string = 'workbook') => {
            const buffer: ArrayBuffer = await workbook.xlsx.writeBuffer()
            const a = document.createElement('a')
            const fileUrl = URL.createObjectURL(new Blob([buffer]))
            a.href = fileUrl
            a.download = `${filename}.xlsx`
            a.click()
            URL.revokeObjectURL(fileUrl)
        },
        workbook,
    }
}
