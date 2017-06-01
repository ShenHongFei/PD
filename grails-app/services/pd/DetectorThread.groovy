package pd

import grails.gorm.transactions.Transactional
import org.hibernate.Hibernate

import java.nio.file.Files
import java.nio.file.StandardCopyOption

import static pd.Application.binDir
import static pd.Application.detectDir
import static pd.DetectorThread.Status.*

@Transactional
class DetectorThread extends Thread{
    
    
    Paper  paper
    Report report
    Student student
    
    Process ps
    // 0=完成 -1=出错 1=正在检测
    Status status
    static enum Status{
        ERROR,FINISHED,RUNNING
    }
    DetectorThread(){}
    DetectorThread(Student student,Paper paper){
        this.paper=paper
        this.student=student
    }
    
    @Override
    void run(){
        status=RUNNING
        def original=paper as File
        def tmp=new File(detectDir,original.name.replaceAll(' ','-')) //论文检测程序不支持带空格的文件名
        def start=System.currentTimeMillis()
        Files.copy(original.toPath(),tmp.toPath(),StandardCopyOption.REPLACE_EXISTING)
        
        String command="$binDir\\PaperFormatDetection.exe $binDir\\temp.docx $tmp false"
        println "论文检测命令行:\n$command"
        ps=command.execute(null as List,binDir)
        ps.waitForOrKill(25*1000)
        
        
        if(ps.exitValue()==0){
            println ps.inputStream.getText('gbk')
            def targetReport = new File(binDir,"Papers/${tmp.name-'.docx'}/report.txt")
            report = new Report(paper,targetReport)
            //这里的student.paper 是被代理的，而且是懒加载，需要重新初始化，才能正常访问papers
            student.papers<<paper
            report.save()
            paper.report=report
            status=FINISHED
            
            targetReport.parentFile.deleteDir()
            tmp.delete()
            println "论文检测正常结束，耗时：${(System.currentTimeMillis()-start)/1000}s"
        }else{
            status=ERROR
            try{
                println ps.errorStream.getText('gbk')
                println ps.inputStream.getText('gbk')
            }catch(Exception e){
                e.printStackTrace()
            }finally{
                println '检测程序出错或killed'
            }
        }
    }
}