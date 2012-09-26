import grails.plugin.jxl.builder.*
class JxlGrailsPlugin {
    // the plugin version
    def version = "0.53"
    // the version or versions of Grails the plugin is designed for
    def grailsVersion = "1.3.6 > *"
    // the other plugins this plugin depends on
    def dependsOn = [:]
    // resources that are excluded from plugin packaging
    def pluginExcludes = [
        "grails-app/views/error.gsp"
    ]

    def title = "Jxl Plugin" // Headline display name of the plugin
    def author = "Shaun Jurgemeyer"
    def authorEmail = "sjurgemeyer@gmail.com"
    def description = '''\
    Plugin for exporting data to Excel using the JXL library.  The plugin allows for formatting as well as formula generation using dynamic methods.
'''

    // URL to the plugin's documentation
    def documentation = "http://grails.org/plugin/jxl"

    // Extra (optional) plugin metadata

    // License: one of 'APACHE', 'GPL2', 'GPL3'
//    def license = "APACHE"

    // Details of company behind the plugin (if there is one)
//    def organization = [ name: "My Company", url: "http://www.my-company.com/" ]

    // Any additional developers beyond the author specified above.
//    def developers = [ [ name: "Joe Bloggs", email: "joe@bloggs.net" ]]

    // Location of the plugin's issue tracker.
    def issueManagement = [ system: "github", url: "https://github.com/sjurgemeyer/grails-jxl-plugin/issues" ]

    // Online location of the plugin's browseable source code.
    def scm = [ url: "https://github.com/sjurgemeyer/grails-jxl-plugin" ]

    def doWithWebDescriptor = { xml ->
        // TODO Implement additions to web.xml (optional), this event occurs before
    }

    def doWithSpring = {
        // TODO Implement runtime spring config (optional)
    }

    def doWithDynamicMethods = { ctx ->
       application.controllerClasses.toList()*.metaClass*.renderExcel= {  Closure closure ->
           def stream = new ByteArrayOutputStream()
           workbook(stream, closure)
           response.contentType = 'application/excel'
           response.outputStream << stream.toByteArray()
        }
    }

    def doWithApplicationContext = { applicationContext ->
        // TODO Implement post initialization spring config (optional)
    }

    def onChange = { event ->
        // TODO Implement code that is executed when any artefact that this plugin is
        // watching is modified and reloaded. The event contains: event.source,
        // event.application, event.manager, event.ctx, and event.plugin.
    }

    def onConfigChange = { event ->
        // TODO Implement code that is executed when the project configuration changes.
        // The event is the same as for 'onChange'.
    }

    def onShutdown = { event ->
        // TODO Implement code that is executed when the application shuts down (optional)
    }
}
