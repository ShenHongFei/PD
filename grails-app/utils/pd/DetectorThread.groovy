package pd

import java.nio.file.Files
import java.nio.file.Paths


class DetectorThread extends Thread{
    
    File binDir
    Paper paper
    File report
    
    Integer waitingSeconds
    
    Process ps
    def exitValue
    
    DetectorThread(File binDir,Paper paper,Integer waitingSeconds){
        this.binDir=binDir
        this.paper=paper
        this.waitingSeconds=waitingSeconds
    }
    
    @Override
    void run(){
        Files.copy(paper.getFile(Application.paperDir),new File(binDir,"detectTmp/$paper.name")
        
        //论文检测程序不支持带空格的文件名
        String command="$binDir/PaperFormatDetection.exe $binDir/temp.docx $paper false"
        println "论文检测命令行:\n$command"
        ps=command.execute(null as List,new File(binDir))
        ps.waitForOrKill(waitingSeconds*1000)
        exitValue=ps.exitValue()
        
        if(0==exitValue){
            println ps.inputStream.getText('gbk')
            new File(reportPath).parentFile.mkdirs()
            Files.move(Paths.get("$BIN_DIR\\Papers\\$paperName\\report.txt"),Paths.get(reportPath))
            new File("$binDir/Papers/$pa").deleteDir()
        }else{
            try{
                println ps.inputStream.getText('gbk')
            }catch(Exception e){
                e.printStackTrace()
            }
            println ps.errorStream.text
            println '检测程序出错或killed'
        }
    }
    
}
