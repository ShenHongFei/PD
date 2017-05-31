package pd

import javax.persistence.Transient
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat

import static pd.Application.reportDir


class Report {
    
    String name
    Paper paper
    static def timeFormat=new SimpleDateFormat('MM-dd-a-h-mm',Locale.CHINA)
    Date createdAt
    
    @Transient
    File reportTmpFile
    
    static constraints = {}
    
    Report(){}
    Report(Paper paper,File reportTmpFile){
        this.paper=paper
        this.reportTmpFile=reportTmpFile
        createdAt=new Date()
        this.name="$paper.name-检测报告-${timeFormat.format(createdAt)}"
    }
    
    
    def afterInsert(){
        Files.move(reportTmpFile.toPath(),(this as File).with{parentFile.mkdirs();toPath()})
    }
    
    def asType(Class type){
        if(type==File){
            return new File(reportDir,"$id/${name}.txt")
        }
    }
}
