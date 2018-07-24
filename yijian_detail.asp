<!--#include file="Conn.asp" -->
<!--#include file="Const1.asp" -->
<!--#include file="Inc/Function.asp" -->
<%set rsm=conn.execute("select * from yijian where id="&request("id")&"")%>
<html>
<head>
<title><%=rsm("title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
  <!--
	p {font-family: "Arial, Helvetica, sans-serif"; font-size: 12pt}
	td {font-family: "Arial, Helvetica, sans-serif"; font-size: 12pt ; line-height:normal; }
	A{font-family: "Arial, Helvetica, sans-serif";text-transform: none; text-decoration: none; font-size: 12pt}
	select{font-family: "Arial, Helvetica, sans-serif";font-size: 12pt;}
	input{font-family: "Arial, Helvetica, sans-serif";font-size: 12pt;}
	a:hover {font-family: "Arial, Helvetica, sans-serif";text-decoration:underline; color: #FFFFFF; font-size: 12pt}
	body {font-family: "Arial, Helvetica, sans-serif";font-size: 12pt}
	div {font-family: "Arial, Helvetica, sans-serif"; font-size: 12pt}
  -->
</style>
<script language="JavaScript1.2" src="js/Function.js"></script>
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
<table width='700' cellpadding='0' cellspacing='0' border='0' style="font-size:18px ">
<tr>		
	<td align="center" height="60">
	<div onClick="window.print();"><font size="+2"><%=rsm("title")%><br><br>
	<%=rsm("addman")%>
	于
	<%=rsm("addtime")%></font></div></td>
</tr>
</table>
<table width='700' cellpadding='0' cellspacing='0' border='1' style="border-collapse:collapse;">
<tr>
	<td colspan="5"><br><br><%=rsm("content")%><br></td>		
</tr>
</table>
<%rsm.close%>
<!--#include file="Index_bottom.asp" -->