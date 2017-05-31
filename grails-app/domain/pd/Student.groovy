package pd

class Student{
    
    String sid //学号
    String email
    String password
    String clazz
    String name
    String gender
    
    
    String  cookieId //**
    String  lastIp //*
    Boolean autologin = false //*
    String  uuid //**
    
    
    static constraints = {
        password size:1..20
        sid matches:/[0-9]{1,20}/,unique:true
        email email:true
    }
    
    @Override
    String toString(){ "{id: $id, sid:$sid, password:$password" }
}