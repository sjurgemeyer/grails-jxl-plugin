grails.project.work.dir = "target"
grails.project.source.level = 1.6
grails.project.target.level = 1.6

grails.project.dependency.resolution = {
    inherits "global"
    log "warn"

    repositories {
        grailsCentral()
    }

    dependencies {
        runtime 'net.sourceforge.jexcelapi:jxl:2.6.12'
    }

    plugins {
        build ":release:2.0.0", {
            export = false
        }
    }
}
