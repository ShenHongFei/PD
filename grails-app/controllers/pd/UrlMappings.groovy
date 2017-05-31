package pd

class UrlMappings {
    
    static mappings = {
        "/**"                           controller:'resource',      action:'resource'
        "/$controller/$action?"{
            /*constraints {
                
            }*/
        }
        "500"(view: '/error')
        "404"(view: '/notFound')
     
        /* 动态匹配action
        "/user/manage/$manageAction"{
            controller='user'
            action={"manage$params.manageAction"}
        }*/
    }
}
