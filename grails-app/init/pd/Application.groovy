package pd

import grails.boot.GrailsApp
import grails.boot.config.GrailsAutoConfiguration
import org.springframework.jdbc.datasource.TransactionAwareDataSourceProxy

import javax.sql.DataSource
import java.text.SimpleDateFormat

class Application extends GrailsAutoConfiguration {
    
    public static def timeFormat=new SimpleDateFormat('yyyy-MM-dd a h:mm',Locale.CHINA)
    public static def fileTimeFormat=new SimpleDateFormat('yyyy-MM-dd-a-h-mm',Locale.CHINA)
    
    public static File projectDir
    public static File webDir
    public static File dataDir
    public static File binDir
    public static File paperDir
    public static File reportDir
    public static File uploadDir
    public static File detectDir
    public static File errorDir

    
//    static Boolean tableExists
    
    static{
        projectDir  =new File(System.properties['user.dir'] as String)
        println "当前路径： $projectDir.absolutePath"
        binDir      =new File(projectDir,'bin')
        webDir      =new File(projectDir,'web')
        (dataDir    =new File(projectDir,'data')).mkdirs()
        (uploadDir  =new File(dataDir,'upload')).mkdirs()
        (paperDir   =new File(dataDir,'paper')).mkdirs()
        (detectDir  =new File(dataDir,'detect')).mkdirs()
        (reportDir  =new File(dataDir,'report')).mkdirs()
        (errorDir   =new File(projectDir,'error')).mkdirs()
    }
    
    @Override
    void doWithApplicationContext(){
        def dataSource=applicationContext.getBean('dataSource') as TransactionAwareDataSourceProxy
        println dataSource.properties
        println config
    }
    
    static void main(String[] args) {
        GrailsApp.run(Application, args)
    }
}


/*        Driver driver=new Driver()
        tableExists = driver.connect("jdbc:h2:file:${System.properties['development']?'D:':'~'}/HM/data/db/HM;AUTO_SERVER=TRUE;MVCC=TRUE;LOCK_TIMEOUT=10000;DB_CLOSE_ON_EXIT=FALSE",[sid:'root'] as Properties).getMetaData().getTables(null,null,"WORD_TYPES",null).next()*/


    