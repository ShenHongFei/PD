import grails.plugin.json.view.JsonViewConfiguration

beans = {
    jsonViewConfiguration(JsonViewConfiguration,{
        prettyPrint=true
    })
}
