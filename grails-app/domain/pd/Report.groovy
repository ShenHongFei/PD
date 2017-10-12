package pd

import javax.persistence.Transient
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat

import static pd.Application.projectDir
import static pd.Application.reportDir
import static pd.Application.reportDir


class Report implements Comparable<Report>{
    
    static constraints = {
        reportTmpFile(nullable: true)
    }
    static transients=['reportTmpFile']
    
    String name
    Date createdAt
    Paper paper
    
    @Transient
    File reportTmpFile
    
    static belongsTo = [paper:Paper]
    static mapping = {
        paper lazy: false
    }
    static fetchMode = [paper:'eager']
    
    Report(){}
    Report(Paper paper,File reportTmpFile){
        this.paper=paper
        this.reportTmpFile=reportTmpFile
        createdAt=new Date()
        this.name="$paper.name-${Application.fileTimeFormat.format(createdAt)}-检测报告"
    }
    
    
    def afterInsert(){
        Files.move(reportTmpFile.toPath(),(this as File).with{parentFile.mkdirs();toPath()})
    }
    
    def asType(Class type){
        if(type==File){
            return new File(reportDir,"$id/${name}.txt")
        }
    }
    
    int compareTo(Report another){
        createdAt<=>another.createdAt
    }
    
    def getDownloadLink(){
        projectDir.toPath().relativize((this as File).toPath()).toString().replaceAll('\\\\','/')
    }
    
    def getHtmlText(){
        new StringBuffer().with{sb->
            (this as File).text.eachLine{sb<<(it+'<br>')}
            sb.toString()
        }
    }
    
    def getCreatedAtString(){
        Application.timeFormat.format(createdAt)
    }
}
