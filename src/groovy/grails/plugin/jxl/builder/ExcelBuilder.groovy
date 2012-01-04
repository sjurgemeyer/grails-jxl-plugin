package grails.plugin.jxl.builder
import grails.plugin.jxl.*

@Mixin(ExcelUtils)
class ExcelBuilder {
    def workbook
    def sheet
    def formula = new ExcelFormulaBuilder()
    def cells = []

    def workbook(String fileName, Closure closure) {
        this.workbook = createWorkbook(fileName)
        closure()
        workbook.write()
        workbook.close()
    }

    def sheet(String name, int position, Closure closure) {
        this.sheet = workbook.createSheet(name, position)
        this.cells = []
        closure()
        cells.each {
            it.write(sheet)
        }
    }

    def cell(col, row, value, props=[:]) {
        def newCell = new Cell(col, row, value, props)
        cells << newCell
        newCell
    }
}
