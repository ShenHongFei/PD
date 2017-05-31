package pd

class MyPage{
    List content
    Integer totalElements
    Integer size
    Integer number //页码 从0计数
    
    
    MyPage(List content,Long totalElements,Integer size,Integer number){
        this.content=content
        this.totalElements=totalElements
        this.size=size
        this.number=number
    }
    
    Integer getTotalPages(){
        Math.ceil(totalElements/size)
    }
    
    Integer getNumberOfElements(){
        last?totalElements-number*size:size
    }
    
    Boolean getFirst(){
        number==0
    }
    
    Boolean getLast(){
        totalPages==0?:number==totalPages-1
    }
}
