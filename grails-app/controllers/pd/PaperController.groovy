package pd

class PaperController {
    
	static responseFormats = ['json']
	
    def count(){
        render Paper.count
    }
    
    @SuppressWarnings("GroovyUnnecessaryReturn")
    def detect(){
        //保存论文
        def paper = new Paper(params.paper,session.student)
        if(!paper.validate()) return render('上传失败，上传的论文有问题')//todo:上传校验
        paper.save(flush:true)
        def detectorThread=session.detectorThread=new DetectorThread(paper)
        detectorThread.start()
        return render('上传成功')
    }
    
    // 0=完成 -1=出错 1=正在检测
    def status(){
        def detectorThread=session.detectorThread
        def status = detectorThread?.status as DetectorThread.Status
        switch(status){
            case null:return render('RUNNING')
            default:return render(status.toString())
        }
    }
    
    
    
}
