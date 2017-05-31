package pd

import org.springframework.core.io.FileSystemResource
import org.springframework.web.accept.PathExtensionContentNegotiationStrategy

import static pd.Application.dataDir
import static pd.Application.webDir

class ResourceController{
    
    static responseFormats = ['json','gsp']
    
    static final HOME_PAGE='/student.html'
    
    def contentNegotiationStrategy=new PathExtensionContentNegotiationStrategy()
    
    def resource(){
        def uri=URLDecoder.decode(request.requestURI-request.contextPath,'UTF-8')
        def resource
        uri-='/ueditor/dialogs/preview'
        if(uri=='/') uri=HOME_PAGE
        if(uri.contains('/data')){
            def fileuri=uri-'/data/'
            println "uri=$fileuri"
            resource=new File(dataDir,fileuri)
        }else{
            println "uri=$uri"
            resource=new File(webDir,uri)
        }
        if(!resource.exists()||resource.directory) {
            println "资源 $uri 不存在"
            return render(view:'/failure',model:[message:"RESOURCE ${uri} NOT FOUND".toString()],status:404)
        }
        response.addHeader('Content-Length',resource.size() as String)
        render(file:resource,contentType:contentNegotiationStrategy.getMediaTypeForResource(new FileSystemResource(resource)))
    }
    
}

