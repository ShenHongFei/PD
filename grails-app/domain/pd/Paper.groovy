package pd

import org.springframework.web.multipart.MultipartFile

import javax.persistence.Transient

import static pd.Application.paperDir

class Paper {
    
    String name
    Student author
    Date uploadAt
    Set<Report> reports=[]
    
    @Transient
    //MultipartFile
    def uploadTmp
    
    static constraints = {}
    
    Paper(){
        
    }
    
    Paper(MultipartFile paper,Student author){
        name=paper.originalFilename-'.docx'
        uploadTmp=paper
        this.author=author
        uploadAt=new Date()
    }
    
    
/*    def getPath(){
        "$author.sid/$filename"
    }
    def getFile(){
        new File(App,path)
    }*/
    
    def getFilename(){
        "${name}.docx"
    }
    
    
    def afterInsert(){
        uploadTmp.transferTo((this as File).with{
            parentFile.mkdirs()
            it
        })
    }
    
    def asType(Class type){
        if(type==File){
            return new File(paperDir,"$id/$filename")
        }else{
            return super.asType(type)
        }
    }
}
