package pd

import grails.gorm.transactions.Transactional
import org.springframework.web.multipart.MultipartFile

import java.nio.file.Files
import java.nio.file.Paths
import java.nio.file.StandardCopyOption

import static pd.Application.binDir
import static pd.Application.detectDir
import static pd.PaperController.DetectorThread.Status.*

@Transactional
class PaperController {
    
	static responseFormats = ['json']
	
    def count(){
        render Paper.count
    }
    
    @SuppressWarnings("GroovyUnnecessaryReturn")
    def detect(){
        def paperFile = params.paper as MultipartFile
        if(!paperFile||paperFile.empty) return render('无上传论文')
        def student = session.student
        def paper = new Paper(paperFile,student)
        if(!paper.validate()) return render('上传失败，上传的论文有问题')//todo:上传校验
        session.paper=paper.save()
        def detectorThread=session.detectorThread=new DetectorThread(student,paper)
        detectorThread.start()
        return render('上传成功')
    }
    
    
    def status(){
        def detectorThread=session.detectorThread
        def status = detectorThread?.status as DetectorThread.Status
        switch(status){
            case null:return render('RUNNING')
            default:return render(status.toString())
        }
    }
    
    //params id(paper)
    def viewReport(){
        def id=params.int('id')
        def report
        if(id){
            report=Paper.get(id)?.report
        }else{
            report = session.paper?.report
        }
        if(!report) return render(view:'/failure',model:[message:'未登录或无检测报告'])
        render(view:'/report/view',model:[report:report])
    }
    
    
    def list(){
        def student = session.student
        if(!student) return render(view:'/failure',model:[message:'用户未登录'])
        def papers = student.papers
        if(!papers||papers.empty) return render(view:'/failure',model:[message:'无检测报告'])
        render(view:'list',model:[papers:papers])
    }
    
    //params pageIndex pageSize
    def listAll(){
        def page        = (params.pageIndex?:0) as Integer
        def size        = (params.pageSize?:5) as Integer
        def sortParams  = (params.sort?:'id,desc').split(',') as List
        def sortBy      = sortParams[0]
        def order       = sortParams[1]
        def papers = Paper.findAll("from Paper as paper inner join fetch paper.author join fetch paper.report order by paper.$sortBy $order".toString(),[max:size,offset:page*size])
        render view:'/mypage',model:[template:'/paper/details',myPage:new MyPage(papers,Paper.count,size,page)]
    }
    
    
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
            //todo:多线程同时访问文件的问题
            status=RUNNING
            def original=paper as File
            def tempPaper=new File(detectDir,original.name.replaceAll(' ','-')) //论文检测程序不支持带空格的文件名
            def tempTemplate=new File(detectDir,"template-${UUID.randomUUID()}.docx")
            def start=System.currentTimeMillis()
            Files.copy(original.toPath(),tempPaper.toPath(),StandardCopyOption.REPLACE_EXISTING)
            Files.copy(Paths.get("$binDir\\template.docx"),tempTemplate.toPath())
            
            String command="$binDir\\PaperFormatDetection.exe $tempTemplate $tempPaper false"
            println "论文检测命令行:\n$command"
            ps=command.execute(null as List,binDir)
            ps.waitForOrKill(25*1000)
            tempTemplate.delete()
            
            if(ps.exitValue()==0){
                println ps.inputStream.getText('gbk')
                def targetReport = new File(binDir,"Papers/${tempPaper.name-'.docx'}/report.txt")
                // todo:！！！！还需要理解一下！！
                Report.withTransaction{
                    report = new Report(paper,targetReport)
                    student.attach()
                    paper.attach()
                    student.papers<<paper
                    paper.report=report
                    report.save(flush:true)
                }
                status=FINISHED
                
                targetReport.parentFile.deleteDir()
                tempPaper.delete()
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
    
}
