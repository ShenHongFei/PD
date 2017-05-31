grails{
    gorm.default.constraints = {
        '*'(nullable: true)
    }
}

eventCompileStart = {
    projectCompiler.srcDirectories << "${basedir}/grails-app/model"
}