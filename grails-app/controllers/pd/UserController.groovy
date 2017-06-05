package pd

import grails.gorm.transactions.Transactional
import groovy.text.SimpleTemplateEngine
import org.springframework.mail.javamail.JavaMailSenderImpl



@Transactional
class UserController {
    
    static responseFormats = ['json']
    
    def templateEngine=new SimpleTemplateEngine()
    
    //1=成功 -1=失败
    @SuppressWarnings("UnnecessaryQualifiedReference")
    def register(){
        def src= applicationContext.getBean('dataSource')
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
            session.student=null
            session.teacher=teacher
            return render(-1)
        }
        if(params.username=='204099999'&&params.password=='123456'){
            if(!Teacher.count) new Teacher(tid:'204099999',name:'赖晓晨',password:'123456').save()
            session.teacher=teacher
            session.student=null
            return render(-1)
        }
        //再尝试学生登录
        def student = Student.find{sid==params.username}
        if(!student) return render(0)
        if(!student.password){
            student.password=params.password
        }
        if(student.password!=params.password) return render(2)
        session.teacher=null
        student.lastIp=request.remoteAddr
        session.student=student
        render(1)
    }
    
    def logout(){
        session.student=null
        session.teacher=null
        redirect(uri:'/login.html')
    }
    
    //当前用户的信息
    def info(){
        render view:'student-info',model:[student:session.student]
    }
    
    def importData(){
        new File('D:/T/file/data.tsv').eachLine{
            def line=it.split('\t')
            new Student(sid:line[0],gender:Student.Gender.valueOf(line[1].toUpperCase()),name:line[3],clazz:line[4],email:line[-1]).save()
        }
        render('success')
    }
    

}
