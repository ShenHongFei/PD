package pd

import org.springframework.web.multipart.MultipartFile

class PaperController {
    
	static responseFormats = ['json']
	
    def count(){
        render Paper.count
    }
    
    @SuppressWarnings("GroovyUnnecessaryReturn")
    def detect(){
        MultipartFile paperFile = params.paper
        if(!paperFile||paperFile.empty) return render('无上传论文')
        def student = session.student
        def paper = new Paper(paperFile,student)
        if(!paper.validate()) return render('上传失败，上传的论文有问题')//todo:上传校验
        session.paper=paper.save(flush:true,failOnError:true)
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
        def id=params.id
        def report
        if(id){
            report=Paper.findById(id)?.report
        }else{
            report = session.paper?.report
        }
        if(!report) return render(view:'/failure',model:[message:'未登录或无检测报告'])
        render(view:'/report/view',model:[report:report])
    }
    
    
    def list(){
        def papers = session.student?.papers
        if(!papers||papers.empty) return render(view:'/failure',model:[message:'未登录或无检测报告'])
        render(view:'list',model:[papers:papers])
    }
    
}
