package grails.plugin.jxl

class ExcelFormulaBuilderTests {
    
    def letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    def testNumberToLetter() {
        assert true
        com.gdb.GdbShell.gdb() 
        def builder = new ExcelFormulaBuilder()
        (0..25).each {
            assert letters[it] == builder.numberToLetter(it)
        }
        (26..51).each {
            assert "A${letters[it-26]}" == builder.numberToLetter(it)
        }
        (260..285).each {
            assert "J${letters[it-260]}" == builder.numberToLetter(it)
        }
    }


}
