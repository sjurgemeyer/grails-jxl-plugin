package grails.plugin.jxl.builder

import grails.plugin.jxl.*
import jxl.format.PageOrientation
import jxl.write.WritableHyperlink

@Mixin(ExcelUtils)
class ExcelBuilder {

    def workbook
    def sheet
    def sheetIndex = 0
    def formula = new ExcelFormulaBuilder()
    def cells = []

    def workbook(String fileName, Closure closure) {
        closure.resolveStrategy =  Closure.DELEGATE_FIRST
        closure.delegate = this
        this.workbook = createWorkbook(fileName)
        sheetIndex = 0
        closure()
        workbook.write()
        workbook.close()
    }

    def workbook(OutputStream stream, Closure closure) {
        closure.resolveStrategy =  Closure.DELEGATE_FIRST
        closure.delegate = this
        this.workbook = createWorkbook(stream)
        sheetIndex = 0
        closure()
        workbook.write()
        workbook.close()
    }

    def sheet(String name="Sheet$sheetIndex", Closure closure) {
        this.sheet = workbook.createSheet(name, sheetIndex++)
        this.cells = []
        closure()
        cells.each {
            it.write(sheet)
        }
    }

    def cell(int col, int row, value, Map props=[:]) {
        def newCell = new Cell(col, row, value, props)
        cells << newCell
        newCell
    }

    def cell(int col, int row, Map props=[:]) {
        def newCell = getCell(sheet, col, row, props)
        cells << newCell
        newCell
    }

    def addData(rowData,startCol=0,startRow=0) {
        addData(sheet, rowData, startCol, startRow)
    }

    def hyperlink(int col, int row, String url) {
        WritableHyperlink writableHyperlink = new jxl.write.WritableHyperlink( col, row, new URL( url ) )
        sheet.addHyperlink(writableHyperlink)
    }

    def setOrientation(PageOrientation pageOrientation) {
        this.sheet.settings.orientation = pageOrientation
    }
}
