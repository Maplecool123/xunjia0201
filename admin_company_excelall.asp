<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> 
<% 
if session("user_id")="" then
	response.write "<script language='javascript'>alert('您还没有登陆，请先登陆');"
	response.write "window.location ='ad_login.asp';</script> "
	Response.End 
end if

if session("iflogin")-2=0 or session("iflogin")-99=0 then
nowfilename="供应商表"&replace(replace(replace(now,":","")," ",""),"/","")
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename = "&nowfilename&".xls"
%>
<html> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<title>供应商表</title>
</head>
<body>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
	<th>ID号</th>
	<th>企业名称</th>
	<th>企业归类</th>
	<th>企业类型</th>
	<th>资质级别</th>
	<th>登陆用户名</th>
	<th>机构代码证号</th>
	<th>税号</th>
	<th>法定代表人</th>
	<th>联系人</th>
	<th>联系电话</th>
	<th>公司地址</th>
	<th>电子邮件</th>
	<th>主营业务</th>
	<th>注册状态</th>
	<th>审核状态</th>
	<th>注册时间</th>
</tr>
<%sql=getsql(request("sql"))
Set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not rs.eof then
	i=0
	do while not rs.eof
	i=i+1%>
<tr align="center">
	<td><%=rs("id")%></td>
	<td><%=rs("companyrealname")%></td>
	<td><%=rs("companyclass")%></td>
	<td><%=shownameint("companytypename","companytype","id",rs("companytype"))%></td>
	<td><%=shownameint("companygradename","companygrade","id",rs("companygrade"))%></td>
	<td><%=rs("companyname")%></td>
	<td><%="&nbsp;"&rs("companyorgno")%></td>
	<td><%="&nbsp;"&rs("companytaxno")&"　"%></td>
	<td><%=rs("companymanager")%></td>
	<td><%=rs("companycontactman")%></td>
	<td><%=rs("companycontacttel")%></td>
	<td><%=rs("companyaddress")%></td>
	<td><%=rs("companyemail")%></td>
	<td><%="&nbsp;"&rs("companymajor")%></td>
	<td><%=getcompanyregstatue(rs("regstatue"))%></td>
	<td><%=getcompanystatue(rs("statue"))%></td>
	<td><%=rs("companyregtime")%></td>
</tr>
	<%Rs.movenext
	loop
else
	response.write "<tr><td colspan='2'><font color='red'>查无资料纪录！</font></td></tr>"
end if

rs.close
set rs = nothing
%>
</table>
</body>
</html>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->