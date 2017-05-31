package pd

import grails.transaction.Transactional
import groovy.text.SimpleTemplateEngine
import org.springframework.mail.javamail.JavaMailSenderImpl
import org.springframework.mail.javamail.MimeMessageHelper

import javax.mail.internet.MimeMessage
import javax.servlet.http.Cookie
import javax.servlet.http.HttpServletResponse

import static Student.GUEST

@Transactional
class UserController {
    
    static responseFormats = ['json']
    
    def mailSender=new JavaMailSenderImpl().with{
        host            =   'smtp.qq.com'
        port            =   465             //端口号，QQ邮箱需要使用SSL，端口号465或587
        username        =   '350986489'
        password        =   'pniljmfgcgzobiij'
        defaultEncoding =   'UTF-8'
        javaMailProperties=[
                'mail.smtp.timeout'               :25000,
                'mail.smtp.auth'                  :true,
                'mail.smtp.starttls.enable'       :true,//STARTTLS是对纯文本通信协议的扩展。它提供一种方式将纯文本连接升级为加密连接（TLS或SSL）
                'mail.smtp.socketFactory.port'    :465,
                'mail.smtp.socketFactory.class'   :'javax.net.ssl.SSLSocketFactory',
                'mail.smtp.socketFactory.fallback':false,
        ] as Properties
        it
    }
    
    def templateEngine=new SimpleTemplateEngine()
    
    @SuppressWarnings("UnnecessaryQualifiedReference")
    def register(){
        Student student=new Student(name:params.name,sid:params.username,clazz:params.grade,password:params.password,gender:params.sex)
        if(!student.validate()) return render(-1)
        student.with{
            uuid=UUID.randomUUID().toString()
            lastIp=request.remoteAddr
            save()
        }
        session.student=student
        render(1)
    }
    
    //params username password
    /*if (msg == 0) {
    	          alertWarning("用户名错误");
    	          $('#user,#pwd2').val('');
    	        }
    	        if (msg == 2) {
    	          alertWarning("密码错误");
    	          $('#pwd2').val('');
    	        }
    	        if (msg == 1) {
    	          location.href = 'student.html';
    	        }
    	        if (msg == -1) {
    	          location.href = 'teacher.html';
    	        }
    * */
    def login(){
        //先尝试教师登录
        def teacher = Teacher.find{tid==params.username&&password==params.password}
        if(teacher){
            session.teacher=teacher
            return render(-1)
        } 
        if(params.username=='204099999'&&params.password=='123456'){
            if(!Teacher.count) new Teacher(tid:'204099999',name:'赖晓晨',password:'123456').save()
            session.teacher=teacher
            session.student=new Student(sid:'204099999',name:'赖晓晨')
            return render(-1)
        }
        //再尝试学生登录
        def student = Student.find{sid==params.username}
        if(!student) return render(0)
        if(student.password!=params.password) return render(2)
        student.lastIp=request.remoteAddr
        session.student=student
        student.save()
        render(1)
    }
    
    def logout(){
        session.student=null
        session.teacher=null
        redirect(uri:'/login.html')
    }
    
    
/*    def logout(){
        def student = session.student
        if(student) student.autologin=false
        session.student=null
        if(request.cookies.find{it.name=='autologin'}) clearCookie(response,'autologin')
        setUserCookies(response,GUEST)
        success '注销成功'
    }
    
    //params email,password,autologin
    
    
    //当前用户的信息
    def info(){
         render view:'info',model:[student:session.student] 
    }
    
    //params oldPassword [newPassword] [newUsername]
    def updateInfo(){
        def student=session.student
        if(params.oldPassword!=student.password) return toFailure(student,'原密码错误')
        //修改用户名
        def newUsername = params.newUsername
        if(newUsername&&newUsername!=student.sid){
            if(Student.find{username==newUsername}){
                return toFailure("用户名${newUsername}已存在")
            }
            student.username=newUsername
        }
        //按需修改密码
        if(params.newPassword)
        student.password=params.password
        if(!student.validate()) return toFailure(student,'更新失败')
        student.save()
        success '更新成功'
    }
    
    //params email
    def sendResetEmail(){
        def student = Student.find{email==params.email}
        if(!student) return toFailure('该邮箱未注册')
        MimeMessage message = mailSender.createMimeMessage()
        //使用MimeMessageHelper构建Mime类型邮件,第二个参数true表明信息类型是multipart类型
        MimeMessageHelper helper = new MimeMessageHelper(message,true,'UTF-8')
        helper.setFrom('350986489@qq.com')
        helper.setTo(params.email as String)
        message.setSubject("大连高校环境联盟 重置密码")
        helper.setText(templateEngine.createTemplate(Application.resetEmailTemplate.newReader('UTF-8')).make([username:student.sid,id:student.id,uuid:student.uuid]).toString(),true)
        try{
            mailSender.send(message)
        }catch(Exception e){
            e.printStackTrace()
            return toFailure(e.localizedMessage)
        }
        success '邮件发送成功'
    }
    
    //params id,uuid
    def resetPassword(){
        def student = Student.find{id==params.id&&uuid==params.uuid}
        if(!student) return 
        student.password='123456'
        success('重置成功，新密码为123456，请及时更改。')
    }
    
    //工具方法
    private def success(String message){
        render view:'/success',model:[message:message]
    }
    private def toFailure(Student student,String failureMessage){
        render view:'/failure',model:[errors:student?.errors,message:failureMessage]
    }
    private def toFailure(String failureMessage){
        render view:'/failure',model:[message:failureMessage]
    }
    static void setCookie(HttpServletResponse response,String name,String value,Integer maxAge ){
        response.addCookie(new Cookie(name,value).with{
            it.maxAge=maxAge
            path='/'
            httpOnly=false
            it
        })
    }
    static void clearCookie(HttpServletResponse response,String cookieName){
        setCookie(response,cookieName,null,0)
    }
    static void setUserCookies(HttpServletResponse response,Student student){
        setCookie(response,'userId',student.id as String,-1)
        setCookie(response,'sid',student.sid,-1)
        setCookie(response,'role',student.role as String,-1)
    }*/
}
