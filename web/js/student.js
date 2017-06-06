var global = (function() { return this || (1,eval)('(this)'); }());
global.fileSize = 0;
window.onload = function(){
    countPaper();
    $('#input01').filestyle({
        buttonText: '浏览'
    });
    loadinfo();
    listPaper();
}

$(document).ready(function() {
    $.ajaxSetup({ cache: false });
    var flag = 0;
    $('#pwd1').change(function() {
        if ($('#pwd1').val().length < 6 || $('#pwd1').val().length > 20 || flag == 1) {
            $('#check1').html("请输入5位以上21位以下密码，且只能是密码或数字~").css('color', 'red').css('font-size', '13px');
        } else {
            $('#check1').html('');
            $('#pwd2').css('border-color', 'rgb(200,200,200)');
        }

    });

    $('#pwd2').change(function() {
        if ($('#pwd2').val().length < 6 || $('#pwd2').val().length > 20) {
            $('#check2').html("请输入5位以上21位以下密码，且只能是密码或数字~").css('color', 'red').css('font-size', '13px');
        } else {
            $('#check2').html('');
            $('#pwd2').css('border-color', 'rgb(200,200,200)');


            $.ajax({
                type: 'post',
                url: 'pass',
                data: {
                    new_pwd: $('#pwd2').val()
                },
                success: function(msg) {
                    if (msg == 1) {
                        $('#check2').html("请输入5位以上21位以下密码，且只能是密码或数字~").css('color', 'red').css('font-size', '13px');
                        flag = 1;
                    } else {
                        $('#check2').html('');
                        $('#pwd2').css('border-color', 'rgb(200,200,200)');
                        flag = 0;
                    }
                },
            });
        }

        if ($('#pwd3').val() != $('#pwd2').val()) {
            $('#check3').html('您两次输入的密码不一致，请重新输入！').css('color', 'red').css('font-size', '13px');
        } else {
            $('#check3').html('');
            $('#pwd3').css('border-color', 'rgb(200,200,200)');
        }
    });

    $('#pwd3').change(function() {
        if ($('#pwd3').val() != $('#pwd2').val()) {
            $('#check3').html('您两次输入的密码不一致，请重新输入！').css('color', 'red').css('font-size', '13px');
        } else {
            $('#check3').html('');
            $('#pwd3').css('border-color', 'rgb(200,200,200)');
        }
    });

    $('#queren').click(function() {
        if ($('#pwd2').val().length < 6 || $('#pwd2').val().length > 20 || flag == 1) {
            $('#check2').html("请输入5位以上21位以下密码，且只能是密码或数字~").css('color', 'red').css('font-size', '13px');
            $('#pwd2,#pwd3').css('border-color', 'red');
            $('#pwd2,#pwd3').val('');
        }
        if ($('#pwd3').val() != $('#pwd2').val()) {
            $('#check3').html('您两次输入的密码不一致，请重新输入！').css('color', 'red').css('font-size', '13px');
            $('#pwd3').css('border-color', 'red');
            $('#pwd3').val('');
        }


        if (flag == 1) {
            $('#pwd2').css('border-color', 'red');
            $('#pwd3').css('border-color', 'red');
            $('#pwd2,#pwd3').val('');
        }


        if ($('#pwd2').val().length < 6 || $('#pwd2').val().length > 20 || $('#pwd3').val() != $('#pwd2').val() || flag == 1) {
            alertWarning('您输入的新密码有问题，请重新输入！');
        } else {
            old = $('#pwd1').val();
            newa = $('#pwd2').val();
            $.ajax({
                data: {
                    password: old,
                    newpassword: newa
                },
                url: 'update',
                success: function(msg) {
                    if (msg == 1) {
                        alertInfo("修改成功");
                        $('#pwd1,#pwd2,#pwd3').val('');
                    } else {
                        alertWarning("您的初始密码有问题！");
                        $('#pwd1').css('border-color', 'red');
                        $('#pwd1').val('');
                    }
                },
            });
        }

    });
});


function loadinfo() {
    $.ajaxSetup({cache:false});
    $.ajax({
        url:'user/info',
        success:function(model){
            if(!model.result){
                console.log(model.message)
                return
            }
            let student=model.student;
            $("#st_name").html(student.name);
            $("#st_id").html(student.sid);
            $("#st_class").html(student.clazz);
            if(student.gender=='MAN'){
                $("#st_sex").html('男');
            }else{
                $("#st_sex").html('女');
            }
        }
    });
}

function countPaper(){
  $.ajaxSetup({ cache: false }); 
  $.ajax({
    type:'get',
    url:'paper/count',
    success:function(msg){
      $('#number').html(msg);
    },
  });
}

//文件大小检测
function fileChange(target,fileSize){
    if(/msie/i.test(navigator.userAgent)&& !window.opera&& !target.files){
        var filePath=target.value;
        var fileSystem=new ActiveXObject("Scripting.FileSystemObject");
        var file=fileSystem.GetFile(filePath);
        fileSize=file.Size;
    }else{
        fileSize=target.files[0].size;
    }
    return fileSize;
}

function listPaper(){
    $.ajaxSetup({cache:false});
    $.ajax({
        url:'paper/list',
        dataType:'json',
        success:function(model){
            if(!model.result){
                console.log(model.message)
                return
            }
            let historyTable=$("#historytable");
            historyTable.empty()
            model.papers.forEach(function(it){
                historyTable.append(`<tr><td>${it.uploadAt}</td><td id="pname_${it.name}">${it.name}</td><td><button type='button' class='btn btn-info' onclick='viewReport(${it.id})'>查看检测报告</td></tr>`)
            })
        }
    })
}


function viewReport(paperId) {
    $.ajaxSetup({ cache: false });
    let data=arguments.length>0?{id:paperId}:null
    $.ajax({
        url: 'paper/view-report',
        data: data,
        success: function(model) {
            var datas = model;
            if(!model.result){
                alertWarning(model.message)
                return
            }
            $("#report-text").empty();
            $('#toolitip1').remove();
            $('#toolitip2').remove();
            $('#download1').remove();
            // $('#download2').remove();
            var n = datas.length;
            $('#report-text').html('<h3 style="text-align:center;">《'+model.report.name+'》</h3>'+'<h5 style="text-align:right;">时间：'+model.report.createdAt+'</h5>'+model.report.content)

            $("#report").append("<a id='download1' href='" + model.report.link + "' class='btn btn-lg btn-primary btn-block' style='margin-top: -17px; width: 90px; margin-left: 921px; display: inline-block !important; font-size: 14px;'>下载报告</a>");
            /*$("#report").append("<a id='download2' href='download?paper_name=" + encodeURI(encodeURI(datas[0])) + "' class='btn btn-lg btn-primary btn-block' style='margin-top: 10px; width: 90px; margin-left: 15px; display: inline-block !important; font-size: 14px;'>自动修正</a>");*/
            $('#report').append('<span id="toolitip1" style="visibility: hidden;width: 134px;font-size: x-small;background-color: rgba(0, 0, 0, 0.54);color: #fff;text-align: center;border-radius: 6px;padding: 6px 0px; /* Position the tooltip */position: absolute;z-index: 1;margin: -47px -114px;">下载论文检测报告</span>')//-220px
            $('#report').append('<span id="toolitip2" style="visibility: hidden;width: 134px;font-size: x-small;background-color: rgba(0, 0, 0, 0.54);color: #fff;text-align: center;border-radius: 6px;padding: 6px 0px; /* Position the tooltip */position: absolute;z-index: 1;margin: -65px -113px;">下载本论文的格式自动修正版本（本功能尚处于测试阶段，对修正结果不做任何承诺）</span>')
            /*$('#download2').mouseenter(function(){
             $('#toolitip2').css('visibility','visible')
             })
             $('#download2').mouseleave(function(){
             $('#toolitip2').css('visibility','hidden')
             })*/
            $('#download1').mouseenter(function(){
                $('#toolitip1').css('visibility','visible')
            })
            $('#download1').mouseleave(function(){
                $('#toolitip1').css('visibility','hidden')
            })
            $('#ptab a:eq(2)').tab('show');
        }
    });
}


function uploadAndDetectPaper(){
    $('#view-report-btn').hide();
    $.ajaxSetup({cache:false});
    listPaper();
    setTimeout("listPaper()","4000");
    var docname=$("#input01").val();
    var docnames=docname.split('.');
    if(docname==""){
        alertWarning("没有选择论文，请重新上传！");
        $('#view-report-btn').hide();
        return;
    }else if(docnames[docnames.length-1]!='docx'){
        alertWarning('论文格式不是.docx，请重新上传！');
        $('#view-report-btn').hide();
        return;
    }

    //传送是否屏蔽代码选项的信息
    /*var p=0;
    if($('#pingbi input').is(':checked'))
        p=1;
    $.ajax({
        type:'post',
        url:'uploadTemp',
        data:{temp:p,},
        success:function(){

        },
    });*/


    //去掉屏蔽代码的提示
    $('#pingbi').addClass('collapse');
    //显示进度条   
    $('#progress,#progresser').removeClass('collapse');
    //连续发送ajax请求，直到得到检测结果，成功：显示查看报告按钮，进度条消失；失败：弹窗提示 
    var delta=10;
    $('#progress p').html('10%');
    $('#progress').css('width','10%');
    var show=setInterval(function(){
        $.ajax({
            type:'get',
            url:'paper/status',
            success:function(msg){
                switch(msg){
                    case 'FINISHED':{
                        $('#progress p').html('100% Complete!');
                        $('#progress').css('width','100%');
                        //延时1s
                        setTimeout(function(){
                            $('#progress,#progresser').addClass('collapse');
                            $('#progress p').html('10%');
                            $('#progress').css('width','10%');
                        },1000);
                        setTimeout(function(){
                            $('#view-report-btn').show();
                            countPaper();
                        },1500);

                        clearInterval(show);
                        break;
                    }
                    case 'ERROR':{
                        alertWarning('很抱歉，检测程序出错，请联系 QQ:957376407')
                        clearInterval(show);
                        break;
                    }
                    case 'RUNNING':{
                        $('#view-report-btn').hide();
                        if(fileSize<26214400){
                            if(delta<90){
                                $('#progress p').html(delta+'%');
                                $('#progress').css('width',delta+'%');
                                delta=delta+10;
                            }else if(delta<99){
                                delta=delta+1;
                                $('#progress p').html(delta+'%');
                                $('#progress').css('width',delta+'%');
                            }
                        }
                        break;
                    }
                }
            },
            error:function(){
                clearInterval(show);
                alertWarning('系统出错，请重新上传！');
            },
        });
    },2000);

    setTimeout(function(){
        $(":file").filestyle('clear');
    },1500);
}


//注销登录
function logout(){
  $.ajax({
    type:'get',
    url:'user/logout',
    success:function(){
    	window.location.href='login.html';
    },
  });
}

function alertWarning(msg) {
  BootstrapDialog.show({
    type: BootstrapDialog.TYPE_DANGER,
    title: "消息提示",
    message: msg,
  });
}

function alertInfo(msg) {
  BootstrapDialog.show({
    type: BootstrapDialog.TYPE_INFO,
    title: "消息提示",
    message: msg,
  });
}
