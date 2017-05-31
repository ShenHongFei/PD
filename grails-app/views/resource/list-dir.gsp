<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <title>槐盟</title>
</head>
<body>
<div>
    <h1>ROOT DIR</h1>
    <g:each in="${projectDir.listFiles()}" var="file">
        <h2><a href="${file.absolutePath}">${file.absolutePath}</a></h2>
    </g:each>
</div>
</body>
</html>

