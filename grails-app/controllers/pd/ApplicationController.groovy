package pd

import grails.core.GrailsApplication
import grails.plugins.*

class ApplicationController implements PluginManagerAware {

    GrailsApplication grailsApplication
    GrailsPluginManager pluginManager
    

    def index() {
        [grailsApplication: grailsApplication, pluginManager: pluginManager]
    }
    
    def refresh(){
        grailsApplication.refresh()
        render 'refresh'
    }
    
    def rebuild(){
        grailsApplication.rebuild()
        render 'rebuild'
    }
    
    def config(){
        def context=applicationContext
        println context
    }
}
