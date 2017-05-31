import grails.plugin.json.view.JsonViewConfiguration
import pd.DetectorThread

beans = {
    jsonViewConfiguration(JsonViewConfiguration,{
        prettyPrint=true
    })
    detectorThread(DetectorThread)
}
