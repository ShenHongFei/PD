package pd

class Advice {
    
    Student advisor
    String content
    Date createdAt

    static constraints = {
        content size:1..500
    }
    
    def beforeInsert(){
        createdAt=new Date()
    }
    
    def getCreatedAtString(){
        Application.timeFormat.format(createdAt)
    }
    
}
