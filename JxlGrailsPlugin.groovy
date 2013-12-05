import grails.plugin.jxl.builder.ExcelBuilder

class JxlGrailsPlugin {
    def version = "0.54"
    def grailsVersion = "1.3.6 > *"
    def title = "Jxl Plugin"
    def author = "Shaun Jurgemeyer"
    def authorEmail = "sjurgemeyer@gmail.com"
    def description = 'Exports data to Excel using the JXL library. The plugin allows for formatting as well as formula generation using dynamic methods.'
    def documentation = "http://grails.org/plugin/jxl"

    def license = "APACHE"
    def issueManagement = [system: "github", url: "https://github.com/sjurgemeyer/grails-jxl-plugin/issues"]
    def scm = [url: "https://github.com/sjurgemeyer/grails-jxl-plugin"]

    def observe = ["controllers"]

    def doWithDynamicMethods = { ctx ->
        application.controllerClasses.toList().each {
            attachRenderExcelMethod( it )
        }
    }

    def onChange = { event -> attachRenderExcelMethod( event.source ) }

    static def attachRenderExcelMethod = { controller ->
        controller.metaClass.renderExcel = { Closure closure ->
            def stream = new ByteArrayOutputStream()
            new ExcelBuilder().workbook( stream, closure )
            response.contentType = 'application/excel'
            response.outputStream << stream.toByteArray()
        }
    }
}
