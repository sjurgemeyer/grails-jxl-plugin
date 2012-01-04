package grails.plugin.jxl
class ExcelFormulaBuilder {
    private static final String letters='ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    static String cell(int colNum, int rowNum) {
        def col = numberToLetter(colNum)
        "$col${rowNum+1}"
    }

    static String range(int colStart, int rowStart, int colEnd, int rowEnd) {
        "${cell(colStart, rowStart)}:${cell(colEnd, rowEnd)}"
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

    def methodMissing(String name, args) {
        "=${name.toUpperCase()}(${args.join(',')})"
    }
}
