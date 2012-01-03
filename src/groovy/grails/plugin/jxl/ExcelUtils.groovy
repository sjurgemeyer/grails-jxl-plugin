package grails.plugin.jxl

import jxl.write.*
import jxl.write.WritableFont.*

class ExcelUtils {

    def addCell(sheet, int columnIndex, int rowIndex, value) {
        sheet.addCell(createCell(columnIndex, rowIndex, value))
    }

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

    public void withTemplateRow(sheet, templateRow, items, Closure closure) {
        def row = templateRow+1
        items.each {
            sheet.insertRow(row)
            closure(it, templateRow, row)
            row++
        }
        sheet.removeRow(templateRow)
    }

    public Cell getCell(sheet, col, row) {
        def jxlCell = sheet.getCell(col,row)
        new Cell(jxlCell)         
    }

    private void eachCell(sheet,cellList,Closure closure) {
        cellList.each { 
            def cell = getCell(sheet, it[0], it[1])
            closure(cell)
            cell.write(sheet)
        }
        
    }

    public void addData(sheet,rowData,startCol=0,startRow=0) {
        rowData.eachWithIndex { row, rowNumber ->
            row.eachWithIndex { col, colNumber -> 
                addFormattedCell(sheet, startCol+colNumber, startRow+rowNumber, col)
            }
        }
    }

    public def createWorkbook(String filePath) {
        jxl.Workbook.createWorkbook(new File(filePath))
    }

    public def createWorkbook(File file) {
        jxl.Workbook.createWorkbook(file)
    }

}
