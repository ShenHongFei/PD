//注销登录
function logout(){
  $.ajax({
    type:'get',
    url:'logout',
    success:function(){
    	window.location.href='login.html';
    },
  });
}
  $(document).ready(function() {
    $('.carousel').carousel({
      interval: 2500
    });
  });