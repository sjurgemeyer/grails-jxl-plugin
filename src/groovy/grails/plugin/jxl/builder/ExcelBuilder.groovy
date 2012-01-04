package grails.plugin.jxl.builder
import grails.plugin.jxl.*

@Mixin(ExcelUtils)
class ExcelBuilder {
    def workbook
    def sheet
    def sheetIndex = 0
    def formula = new ExcelFormulaBuilder()
    def cells = []

    def workbook(String fileName, Closure closure) {
        this.workbook = createWorkbook(fileName)
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
        println "first method $col $row"
        def newCell = new Cell(col, row, value, props)
        cells << newCell
        newCell
    }

    def cell(int col, int row, Map props=[:]) {
        println "second method"
        def newCell = getCell(sheet, col, row, props)
        cells << newCell
        newCell
    }

    def addData(rowData,startCol=0,startRow=0) {
        addData(sheet, rowData, startCol, startRow)
    }
}
