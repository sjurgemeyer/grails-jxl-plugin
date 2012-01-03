package grails.plugin.jxl
class ExcelFormulaBuilder {
    static final String letters='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    static String sum(colNum, startRow, endRow) {
        def col = numberToLetter(colNum)
        return "=SUM($colNum$startRow:$colNum$endRow)" 
    }

    static String numberToLetter(int num) {
       num = num+1
       String result = ''
       while (num > 0) {
           int remainder = num%26
           result = letters[remainder-1] + result
           num = (num-1)/26
       }
       result
    }
}
