package pd

class Advice {
    
    Student advisor
    String text
    Date createdAt

    static constraints = {
    }
    
    def beforeCreate(){
        createdAt=new Date()
    }
    
}
