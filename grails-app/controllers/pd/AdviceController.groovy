package pd
import grails.gorm.transactions.Transactional


@Transactional
class AdviceController {
	static responseFormats = ['json']
	
    def submit(){
        def student = session.student
        new Advice(content:params.content,advisor:student).save()
        render(1)
    }
    
    def list(){
        def page        = (params.pageIndex?:0) as Integer
//        def size        = (params.pageSize?:5) as Integer
        def size        = 1000
        def sortParams  = (params.sort?:'id,desc').split(',') as List
        def sortBy      = sortParams[0]
        def order       = sortParams[1]
        def advices = Advice.findAll("from Advice as advice inner join fetch advice.advisor order by advice.$sortBy $order".toString(),[max:size,offset:page*size])
        render view:'/mypage',model:[myPage:new MyPage(advices,Advice.count,size,page), template:'/advice/details']
    }
}
