e:
cd E:\SDK\PD\
ls

# 单机程序编译后添加运行依赖（资源文件）
    cd E:\SDK\PD\app\postgraduate\PaperFormatDetection
    ls .\bin\Debug
    ls .\res
    copy .\res\* .\bin\Debug\

# 测试单机程序
    cd .\bin\Debug\
    # 研究生命令行
    chcp 936
    E:\SDK\PD\app\postgraduate\PaperFormatDetection\bin\Debug\PaperFormatDetection.exe E:\SDK\PD\app\postgraduate\PaperFormatDetection\bin\Debug\temp.docx E:\SDK\PD\test-papers\post-graduate\（陈瑞鑫）基于SWOT分析的企业财务管理系统设计与开发.docx false 1 false
    chcp 65001
    # 查看报告
    ls .\Papers -Recurse
    
    # 清理测试
    rm .\Templates,.\Papers -Recurse
  

# make release
    cd E:\SDK\PD\
    ls release
    # 构建 后端jar包
        .\gradlew.bat bootrepackage

    # 复制
        mkdir .\release -ErrorAction SilentlyContinue
        cp run-*.ps1 .\release
        cp .\教师账号.txt .\release
        cp .\build\libs\PD.jar .\release
        cp .\web .\release  -Recurse -Force
        mkdir .\release\app\postgraduate\PaperFormatDetection\bin\ -ErrorAction SilentlyContinue
        cp .\app\postgraduate\PaperFormatDetection\bin\Debug .\release\app\postgraduate\PaperFormatDetection\bin\Debug -Recurse -Force

    # 检查
        ls .\release
        ls .\release\web\
        ls .\release\app\postgraduate\PaperFormatDetection\bin\Debug\

    # 打包
        Add-Type -AssemblyName "System.IO.Compression.FileSystem"
        [IO.Compression.ZipFile]::CreateFromDirectory("$PWD\release","$pwd\PaperDetect-2017-10-12.zip")

    # 清理
        rm PaperDetect-*.zip
        rm .\release -Recurse