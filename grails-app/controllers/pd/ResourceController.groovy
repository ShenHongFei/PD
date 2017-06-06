package pd

import org.springframework.core.io.FileSystemResource
import org.springframework.web.accept.PathExtensionContentNegotiationStrategy

import static pd.Application.dataDir
import static pd.Application.webDir

class ResourceController{
    
    static responseFormats = ['json','gsp']
    
    def contentNegotiationStrategy=new PathExtensionContentNegotiationStrategy()
    
    def resource(){
        def uri=URLDecoder.decode(request.requestURI-request.contextPath,'UTF-8')
        def resource
        if(uri=='/'){
            println "URI=\t$uri"
            if(session.teacher) uri='/teacher.html'
            else if(session.student) uri='/student.html'
            else uri='/login.html'
        }
        if(uri.contains('/data')){
            def dataUri=uri-'/data/'
            println "DATA-URI=\t$dataUri"
            resource=new File(dataDir,dataUri)
        }else{
            println "WEB-URI=\t$uri"
            resource=new File(webDir,uri)
        }
        if(!resource.exists()||resource.directory) {
            println "资源 $uri 不存在"
            return render(view:'/failure',model:[message:"RESOURCE ${uri} NOT FOUND".toString()],status:404)
        }
        response.addHeader('Content-Length',resource.size() as String)
        if(uri.endsWith('.txt')||uri.endsWith('.docx')){
            response.addHeader('Content-Disposition',"attachment; filename=\"${URLEncoder.encode(resource.name,'UTF-8')}\"")
        }
        try{
            return render(file:resource,contentType:contentNegotiationStrategy.getMediaTypeForResource(new FileSystemResource(resource)))
        }catch(any){}
    }
    
}

