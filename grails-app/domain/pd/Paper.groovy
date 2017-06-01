package pd

import org.apache.commons.collections.comparators.ComparableComparator
import org.springframework.web.multipart.MultipartFile

import javax.persistence.Transient
import java.text.SimpleDateFormat

import static pd.Application.paperDir

class Paper implements Comparable<Paper>{
    
    String name
    Student author
    Date uploadAt
    Report report
    
    @Transient
    //MultipartFile
    def uploadTmp
    
    public static def timeFormat=new SimpleDateFormat('MM-dd-a-h-mm',Locale.CHINA)

    static mappings={
        author lazy:false
    }
    
/*    static constraints = {}*/
    
   
    
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
    
    int compareTo(Paper another){
        uploadAt<=>uploadAt
    }
}
