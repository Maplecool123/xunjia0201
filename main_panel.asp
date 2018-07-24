<!--#include file="conn.asp" -->
<!--#include file="const.asp" -->
<!--#include file="Inc/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<script language="JavaScript" type="text/JavaScript">
function logout()
{
if (confirm ("您真的要登出系统吗?"))
	top.location.replace('Index_logout.asp');
}
</script>
<style type="text/css">
<!--
p {font-family: "Arial, Helvetica, sans-serif"; font-size: 10pt}
td {font-family: "Arial, Helvetica, sans-serif"; font-size: 10pt ; line-height:normal; }
A{font-family: "Arial, Helvetica, sans-serif";text-transform: none; text-decoration: none; font-size: 10pt}
select{font-family: "Arial, Helvetica, sans-serif";font-size: 10pt;}
input{font-family: "Arial, Helvetica, sans-serif";font-size: 10pt;}
a:hover {font-family: "Arial, Helvetica, sans-serif";text-decoration:underline; color: #FFFFFF; font-size: 10pt}
body {font-family: "Arial, Helvetica, sans-serif";font-size: 10pt}
div {font-family: "Arial, Helvetica, sans-serif"; font-size: 10pt}
-->
</style>
<link href='css/panpel.css' rel='stylesheet' type='text/css'>
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="main">
	<div class="logo"></div>
    	<div class="comm">
		<button class="set_homepage" onClick="this.style.behavior='url(#default#homepage)';this.setHomePage(top.location.href);">设为首页</button>&nbsp;&nbsp;
      	<button class="mainpage" onClick="top.main.location.href='main_main.asp'">主页</button>&nbsp;&nbsp;
      	<button class="exit" onClick="logout();">登出</button>
	</div>
	<br>
	<font class="info">
	<img src= "images/cur_shop.png" align="absmiddle"></img>
	<font class='name'>您是：<%=session("loginas")%></font>     &nbsp;&nbsp;&nbsp;
	<img src= "images/cur_user.png" align="absmiddle"></img>
	<font class='name'><%=session("userrealname")%></font> / <font class='user'><%=session("username")%></font>
	</font>     &nbsp;&nbsp;&nbsp;
	<marquee direction="left" onMouseOver="this.stop();this.style.cursor='pointer';" onMouseOut="this.start();" width="300" scrollamount="2"><%=gonggao%></marquee>
</div>
</body>
</html>