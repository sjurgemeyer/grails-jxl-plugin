package grails.plugin.jxl

import grails.plugin.jxl.builder.ExcelFormulaBuilder

class ExcelFormulaBuilderTests extends GroovyTestCase {

    private letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

	 private ExcelFormulaBuilder builder = new ExcelFormulaBuilder()

    void testNumberToLetter() {
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

    void testDynamicMethods() {
        assert "=SUM(A1:A5)" == builder.sum(builder.range(0,0,0,4))
    }
}
