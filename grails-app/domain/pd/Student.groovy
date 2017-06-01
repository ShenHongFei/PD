package pd

import javax.persistence.FetchType
import javax.persistence.OneToMany

class Student{
    
    String sid //学号
    String email
    String password
    String clazz
    String name
    String gender
    
    List<Paper> papers=[]

    
    
    String  cookieId //**
    String  lastIp //*
    Boolean autologin = false //*
    String  uuid //**
    
    static constraints = {
        password size:1..20
        sid matches:/[0-9]{1,20}/,unique:true
        email email:true
    }
    static hasMany = [papers:Paper]
    static mapping={
        papers lazy:false
    }
    

    
    @Override
    String toString(){ "{id: $id, sid:$sid, password:$password" }
}