package pd

import grails.boot.GrailsApp
import grails.boot.config.GrailsAutoConfiguration

class Application extends GrailsAutoConfiguration {
    
    public static File projectDir
    public static File webDir
    public static File dataDir
    public static File binDir
    public static File paperDir
    public static File uploadTmpDir


    
//    static Boolean tableExists
    
    static{
        projectDir=new File(System.properties['user.dir'] as String)
        println "当前路径： $projectDir.absolutePath"
        
        webDir=new File(projectDir,'web')
    
        (dataDir=new File(projectDir,'data')).mkdirs()
        (uploadTmpDir=new File(dataDir,'uploadTmp')).mkdirs()
        (paperDir=new File(dataDir,'paper')).mkdirs()
        
        binDir=new File(projectDir,'bin')
    }
    
    @Override
    void doWithApplicationContext(){
        println config
    }
    
    static void main(String[] args) {
        GrailsApp.run(Application, args)
    }
}


/*        Driver driver=new Driver()
        tableExists = driver.connect("jdbc:h2:file:${System.properties['development']?'D:':'~'}/HM/data/db/HM;AUTO_SERVER=TRUE;MVCC=TRUE;LOCK_TIMEOUT=10000;DB_CLOSE_ON_EXIT=FALSE",[sid:'root'] as Properties).getMetaData().getTables(null,null,"WORD_TYPES",null).next()*/


    