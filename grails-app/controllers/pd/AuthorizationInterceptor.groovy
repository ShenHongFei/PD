package pd

import org.hibernate.SessionFactory

class AuthorizationInterceptor {
    
    SessionFactory sessionFactory
    
    AuthorizationInterceptor(){
        matchAll()
    }

    boolean before() {
        def uri = URLDecoder.decode(request.requestURI-request.contextPath,'UTF-8')
        println "原始请求： ${uri}"
        def filterList=['/student.html','/teacher.html','/']
//        println session.student
//        println session.teacher
        if(filterList.contains(uri)&&!session.student&&!session.teacher) return redirect(uri:'/login.html')
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
//        sessionFactory.currentSession.flush()
        true
    }

    void afterView() {
        if(request.getHeader('Origin')){
            response.addHeader('Access-Control-Allow-Origin','shenhongfei.site')
            response.addHeader('Access-Control-Allow-Credentials','true')
        }
    }
}
