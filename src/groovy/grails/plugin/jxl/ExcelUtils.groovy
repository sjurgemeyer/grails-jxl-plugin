package grails.plugin.jxl

import jxl.write.*
import jxl.write.WritableFont.*

class ExcelUtils {

    void copyCellFormatWithValue(sheet, origCol, origRow, newCol, newRow, newValue) {

        def format = sheet.getCell(origCol,origRow).cellFormat
        Cell cell = new Cell(newCol, newRow, newValue)
        if (format) {
            cell.format = new WritableCellFormat(format)
            cell.font = new WritableFont(format.font)
        }
        cell.write(sheet)
    }

    void copyDown(sheet, origCol, origRow, newRow) {
        sheet.addCell(sheet.getCell(origCol,origRow).copyTo(origCol, newRow))
    }

    void copyDown(sheet, origCol, origRow, newRow, newValue) {
        copyCellFormatWithValue(sheet, origCol, origRow, origCol, newRow, newValue)
    }

    void mergeAcross(sheet, startCol, endCol, row) {
       sheet.mergeCells(startCol, row, endCol, row)
    }

    void autosizeColumn(sheet, int col, value = true) {
        def columnView = sheet.getColumnView(col)
        columnView.setAutosize(value)

        sheet.setColumnView(col, columnView)
    }

    void autosizeColumn(sheet, Range columnRange, value = true) {
        columnRange.each { int col ->
            autosizeColumn(sheet, col, value)
        }
    }

    void withTemplateRow(sheet, templateRow, items, Closure closure) {
        def row = templateRow+1
        items.each {
            sheet.insertRow(row)
            closure(it, templateRow, row)
            row++
        }
        sheet.removeRow(templateRow)
    }

    Cell getCell(sheet, int col, int row, Map props=[:]) {
        def jxlCell = sheet.getCell(col,row)
        new Cell(jxlCell, props)
    }

    private void eachCell(sheet,cellList,Closure closure) {
        cellList.each {
            def cell = getCell(sheet, it[0], it[1])
            closure(cell)
            cell.write(sheet)
        }
    }

    void addData(sheet,rowData,startCol=0,startRow=0) {
        rowData.eachWithIndex { row, rowNumber ->
            row.eachWithIndex { col, colNumber ->
                new Cell(startCol+colNumber, startRow+rowNumber, col).write(sheet)
            }
        }
    }

    def createWorkbook(String filePath) {
        jxl.Workbook.createWorkbook(new File(filePath))
    }

    def createWorkbook(File file) {
        jxl.Workbook.createWorkbook(file)
    }

    def createWorkbook(OutputStream stream) {
        jxl.Workbook.createWorkbook(stream)
    }
}
