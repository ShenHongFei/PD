package pd


import grails.rest.*
import grails.converters.*
import org.springframework.web.multipart.MultipartFile

import static pd.Application.paperDir

class PaperController {
	static responseFormats = ['json', 'xml']
	
    def count(){
        render Paper.count
    }
    
    def detect(){
        //保存论文
        def paper = new Paper(params.paper,session.student).with{save()}
        
        
        String command="$BIN_DIR\\PaperFormatDetection.exe $BIN_DIR\\temp.docx $paperPath false"
        
        println paper.originalFilename
    }
}
