<!--#include file="conn.asp" -->
<!--#include file="const.asp" -->
<!--#include file="Inc/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style type="text/css">
<!--
.main {width: 100%;height: 30px;background-image:url(images/title_bg.png);}
.content {width: 100%;height: 20px;margin-top:6px;}
.info {
 color: #CCCCCC;
 font-family:Arial, Helvetica, sans-serif;
 font-size:10pt;
}
.name {
 color:#FFFFFF;
 font-family:Arial, Helvetica, sans-serif;
 font-size:10pt;
}
-->
</style>
</head>
<body bgcolor="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class="main">
<div class="content">
	<center>
	<font class="info">
	<img src= "images/country.png" align="absmiddle"></img>
	<font class='name'>联系方式：<%=sysmanagercontact%></font>&nbsp;&nbsp;&nbsp;&nbsp;
	<img src= "images/kugar.png" align="absmiddle"></img>
	<font class='name'>联系电话：<%=syscontactphone%></font>&nbsp;&nbsp;&nbsp;&nbsp;
	<img src= "images/cur_user.png" align="absmiddle"></img>
	<font class='name'>负责人：<%=sysmanager%></font>
	</font>
	</center>
</div>
</div>
</body>
</html>