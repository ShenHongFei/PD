package pd

import org.springframework.web.multipart.MultipartFile

import javax.persistence.Transient

import static pd.Application.paperDir

class Paper {
    
    
    String name
    String filename
    @Transient
    File file
    Student author
    Date uploadAt
    
    Paper(){
        
    }
    
    Paper(MultipartFile paper,Student author){
        paper.transferTo(new File(paperDir,"$author.sid/$paper.originalFilename"))
        this.filename=filename
        this.name=filename-'.docx'
        this.author=author
        this.file=new File(paperDir,)
    }
    
    
    def getPath(){
        "$author.sid/$filename"
    }
    def getFile(){
        new File(App,path)
    }


    static constraints = {
    }
    
    def beforeCreate(){
        uploadAt=new Date()
    }
}
