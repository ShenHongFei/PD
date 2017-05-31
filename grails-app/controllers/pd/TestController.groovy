package pd

class TestController {
	static responseFormats = ['json', 'xml']
	
    def index() {
        def user = new Student(id:100,email:'aaa@bbb.com',password:'123').save()
        println user.id
        render user.id
        
    }
}
