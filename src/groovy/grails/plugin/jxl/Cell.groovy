package grails.plugin.jxl

import static jxl.write.WritableFont.*
import jxl.*
import jxl.write.WritableCellFormat
import jxl.write.WritableFont

class Cell {
    Class clazz
    def jxlCell
    WritableFont font = new WritableFont(ARIAL)
    WritableCellFormat format = new WritableCellFormat(font)

    def methodMissing(String name, args) {
        if (font.metaClass.respondsTo(font, name, args)) {
            return font.invokeMethod(name, args)
        }
        if (format.metaClass.respondsTo(format, name, args)) {
            return format.invokeMethod(name, args)
        }
        if (name.endsWith('Border')) {
            borderMethod(name - 'Border')
        } else throw new MissingMethodException(name, delegate, args)
    }

    def propertyMissing(String name) {
        if (jxlCell.hasProperty(name)) return jxlCell."$name"
        throw new MissingPropertyException(name)
    }

    def propertyMissing(String name, value) {
        if (font.hasProperty(name)) font."$name" = value
        else if (format.hasProperty(name)) format."$name" = value
        else throw new MissingPropertyException(name)
    }

    Cell (int columnIndex, int rowIndex, Number value, Map props=[:]) {
        jxlCell = new jxl.write.Number(columnIndex, rowIndex, value)
        this.properties = props
    }

    Cell (int columnIndex, int rowIndex, Boolean value, Map props=[:]) {
        jxlCell = new jxl.write.Boolean(columnIndex, rowIndex, value)
        this.properties = props
    }

    Cell (int columnIndex, int rowIndex, Date value, Map props=[:]) {
        jxlCell = new jxl.write.DateTime(columnIndex, rowIndex, value)
        this.properties = props
    }

    Cell (int columnIndex, int rowIndex, Object value, Map props=[:]) {
        jxlCell = new jxl.write.Blank(columnIndex, rowIndex)
        this.properties = props
    }

    Cell (int columnIndex, int rowIndex, String value, Map props=[:]) {
        if (value && value[0] == '=') {
            jxlCell = new jxl.write.Formula(columnIndex, rowIndex, value[1..-1])
        } else {
            jxlCell = new jxl.write.Label(columnIndex, rowIndex, value)
        }
        this.properties = props
    }

    Cell (jxl.write.biff.CellValue existingCell, Map props=[:]) {

        jxlCell =createJxlCell(existingCell)
        font = existingCell?.cellFormat?.font?new WritableFont(existingCell.cellFormat.font):new WritableFont(ARIAL)
        format = existingCell?.cellFormat?new WritableCellFormat(existingCell.cellFormat):new WritableCellFormat(font)
        jxlCell.cellFormat = format
        this.properties = props
    }

    Cell (jxl.biff.EmptyCell existingCell, Map props=[:]) {

        jxlCell = new jxl.write.Label(existingCell.col, existingCell.row, "")
        font = existingCell?.cellFormat?.font?new WritableFont(existingCell.cellFormat.font):new WritableFont(ARIAL)
        format = existingCell?.cellFormat?new WritableCellFormat(existingCell.cellFormat):new WritableCellFormat(font)
        jxlCell.cellFormat = format
        this.properties = props
    }

    private borderMethod(String borderName) {
        setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle."${toStaticField(borderName)}")
        this
    }

    private toStaticField(String s) {
        s.replaceAll(/([A-Z])/) { "_${it[0]}" }.toUpperCase()
    }

    private createJxlCell(jxl.write.biff.CellValue cell) {
        return cell.class.newInstance(cell)
    }

    private void setProperties(Map props) {
        if (props.font) {
            this.font = new WritableFont(new WritableFont.FontName(props.font))
        }
        props.each { key, value ->
            if (key != 'font') this."$key" = value
        }
    }

    void write(Sheet sheet) {
       format.font = font
       jxlCell.setCellFormat(format)
       sheet.addCell(jxlCell)
    }

    Cell bold() {
        font.boldStyle = WritableFont.BOLD
        this
    }

    Cell italic() {
        font.italic = true
        this
    }

    Cell center() {
        format.alignment = jxl.format.Alignment.CENTRE
        this
    }

    Cell centre() {
        format.alignment = jxl.format.Alignment.CENTRE
        this
    }

    Cell left() {
        format.alignment = jxl.format.Alignment.LEFT
        this
    }

    Cell right() {
        format.alignment = jxl.format.Alignment.RIGHT
        this
    }

    Cell fill() {
        format.alignment = jxl.format.Alignment.FILL
        this
    }

    Cell justify() {
        format.alignment = jxl.format.Alignment.JUSTIFY
        this
    }

    Cell wrap() {
        format.wrap = true
        this
    }
}
