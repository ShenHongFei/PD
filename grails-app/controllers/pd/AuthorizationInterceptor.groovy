package pd

import org.hibernate.SessionFactory

class AuthorizationInterceptor {
    
    SessionFactory sessionFactory
    
    AuthorizationInterceptor(){
        matchAll()
    }

    boolean before() {
        def uri = URLDecoder.decode(request.requestURI-request.contextPath,'UTF-8')
        try{
            if(!uri.split('/')[-1].contains('.')){
                println "API=\t$uri"
            }
        }catch(any){}
        if(uri.endsWith('.jsp')) return redirect(uri:'/')
        def filterExcludeList=['/login.html','/register.html']
        if(!filterExcludeList.contains(uri)&&uri.endsWith('.html')&&!session.student&&!session.teacher) return redirect(uri:'/login.html')
        if(session.teacher&&uri=='/student.html') return redirect(uri:'/teacher.html')
        if(session.student&&uri=='/teacher.html') return redirect(uri:'/student.html')
        if((session.teacher||session.student)&&uri=='/login.html') return redirect(uri:'/')
        /*def refreshCookie=false
        //首次访问网站
        if(!session.user){
            session.user=GUEST
            refreshCookie=true
        }
        //若当前用户为访客，根据有无autologinCookie尝试自动登录
        def autologinCookie = request.cookies.find{it.name=='autologin'}
        if(session.user==GUEST&&autologinCookie){
            def user = Student.find{lastIp==request.remoteAddr&&cookieId==autologinCookie.value&&autologin}
            if(user){
                session.user=user
                refreshCookie=true
            }else{
                UserController.clearCookie(response,'autologin')
            }
        }
        //设置/刷新 前端Cookie
        if(refreshCookie||!request.cookies.find{it.name=='userId'}){
            UserController.setUserCookies(response,session.user)
        }
        true*/
        
        true
    }
    
    boolean after() {
        true
    }

    void afterView() {
        
    }
}
