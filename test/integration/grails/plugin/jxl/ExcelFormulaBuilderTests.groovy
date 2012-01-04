package grails.plugin.jxl

import static ExcelFormulaBuilder.*
class ExcelFormulaBuilderTests {
    
    def letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    def testNumberToLetter() {
        assert true
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


    def testDynamicMethods() {
        def builder = new ExcelFormulaBuilder()
        assert "=SUM(A1:A5)" == builder.sum(range(0,1,0,5))
    }


}
