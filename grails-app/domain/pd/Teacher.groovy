package pd

class Teacher {
    
    String tid
    String password
    String name
    String email
    

    static constraints = {
        password size:1..20
        tid matches:/[0-9]{1,20}/,unique:true
        email email:true
    }
    
    @Override
    String toString(){ "{id: $id, tid:$tid, password:$password" }
}
