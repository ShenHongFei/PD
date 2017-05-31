//自动登录问题
var cookieName_username = "LOGIN_USER_NAME_TEST";
var cookieName_password = "LOGIN_PASSWORD_TEST";
var cookieName_autologin = "LOGIN_AUTO_TEST";

//得到Cookie信息
function getUserInfoByCookie() {
  var uname = getCookie(cookieName_username);
  if (uname != null && !uname.toString().isnullorempty()) {
    GetObj('user').value = uname;
  }

  var upass = getCookie(cookieName_password);
  if (upass != null && !upass.toString().isnullorempty()) {
    GetObj('pwd2').type = 'password';
    GetObj('pwd2').value = upass;
  }
  var autologin = getCookie(cookieName_autologin);
  if (autologin != null && !autologin.toString().isnullorempty())
    if (autologin.toString().trim() == "true") {
      GetObj('ck_rmbUser').checked = true;
    } else {
      GetObj('ck_rmbUser').checked = false;
    }
}

$(document).ready(function() {
  $('#forget').click(function() {
    if ($('#user').val() == '201612345') {
      alertWarning('您不能进行密码验证，请联系管理员重置密码！');
    } else {
      window.location.href = 'security_01.html';
    }
  });
  $('#user').bind('change', function() {
    //alertWarning((/[2-4](08|09|10|11|12|13|14|15)\d{5}/).test($('#user').val()));
    if($("#user").val().length == 0){

      $("#user").attr("data-content","这里不可以空着哦~");
      $('#user').popover('show');
    } else if(!(/20(0|1|2|3|4)\d{6}/).test($('#user').val())) {
      $('#user').attr('data-content',"您的学号格式不对哦~");
      $('#user').popover('show');
    } else {

      $('#user').popover('hide');
      $('#user').css('border-color', 'rgb(200,200,200)');
    }
  });


  $("#ck_rmbUser").bind("click", function () {
    if(!$('#ck_rmbUser').checked){
      delCookie(cookieName_username);
      delCookie(cookieName_password);
      delCookie(cookieName_autologin);
  }

  });
  $('#pwd2').bind('change', function() {
    if ($('#pwd2').val().length == 0) {
      $("#pwd2").attr("data-content","这里不可以空着哦~");
      $('#pwd2').popover('show');
    }else if(!(/[0-9a-zA-Z]{5,20}/).test($('#pwd2').val())){
      $('#pwd2').attr('data-content',"您的密码位数或符号不对哦~");
      $('#pwd2').popover('show');
    } else {
      $('#pwd2').popover('hide');
      $('#pwd2').css('border-color', 'rgb(200,200,200)');
    }
  });
  
  $('#denglu').click(function() {

      if(($('#user').val().length == 0)||($('#pwd2').val().length == 0)){
        if ($('#user').val().length == 0) {
          $('#user').css('border-color', 'red');
        }
         if ($('#pwd2').val().length == 0) {
          $('#pwd2').css('border-color', 'red');
        }
        alertWarning("内容不可以为空！");
        return;
      }else if(!(/20(0|1|2|3|4)\d{6}/).test($('#user').val())||!(/[0-9a-zA-Z]{5,20}/).test($('#pwd2').val())){
        if(!(/20(0|1|2|3|4)\d{6}/).test($('#user').val())){
          $('#user').css('border-color', 'red');
        }
        if (!(/[0-9a-zA-Z]{5,20}/).test($('#pwd2').val())) {
          $('#pwd2').css('border-color', 'red');
        }
        alertWarning("您的输入有问题！");
        return;
      }
      
      $.ajax({
          type: 'post',
          url: 'user/login',
          //asyc:false,
          data: {
            username: $('#user').val(),
            password: $('#pwd2').val()
          },
          success: function(msg) {
            if (msg == 0) {
              alertWarning("用户名错误");
              $('#user,#pwd2').val('');
            }
            if (msg == 2) {
              alertWarning("密码错误");
              $('#pwd2').val('');
            }
            if (msg == 1) {
              location.href = 'student.html';
            }
            if (msg == -1) {
              location.href = 'teacher.html';
            }
          },
        });
    });
    //取消以前的信息
    delCookie(cookieName_username);
    delCookie(cookieName_password);
    var autologin = GetObj('ck_rmbUser');
    //保存在新的cookie中
    if (autologin.checked) {
      SetCookie(cookieName_username, $("#user").val(), 7);
      SetCookie(cookieName_password, $("#pwd2").val(), 7);
    }
    SetCookie(cookieName_autologin, autologin.checked, 7);
    if(($('#user').val().length != 0) &&($('#pwd2').val().length !=0 )){
    	 $.ajax({
    	      type: 'post',
    	      url: 'user/login',
    	      //asyc:false,
    	      data: {
    	        username: $('#user').val(),
    	        password: $('#pwd2').val()
    	      },
    	      success: function(msg) {
    	        if (msg == 0) {
    	          alertWarning("用户名错误");
    	          $('#user,#pwd2').val('');
    	        }
    	        if (msg == 2) {
    	          alertWarning("密码错误");
    	          $('#pwd2').val('');
    	        }
    	        if (msg == 1) {
    	          location.href = 'student.html';
    	        }
    	        if (msg == -1) {
    	          location.href = 'teacher.html';
    	        }
    	      },
    	    });
    }
   
  });