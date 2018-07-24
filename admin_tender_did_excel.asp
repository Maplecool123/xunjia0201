<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> 
<% 
if session("user_id")="" then
	response.write "<script language='javascript'>alert('您还没有登陆，请先登陆');"
	response.write "window.location ='ad_login.asp';</script> "
	Response.End 
end if

if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
nowfilename="物资汇总表"&replace(replace(replace(now,":","")," ",""),"/","")
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename = "&nowfilename&".xls"
%><html> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<title>物资汇总表</title>
</head>
<body>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
	<th>ID号</th>
	<th>项目名称</th>
	<th>竞价编号</th>
	<th>采购依据</th>
	<th>项目类型</th>
	<th>竞价等级</th>
	<th>需求日期</th>
	<th>竞价时间</th>
	<th>发布人</th>
	<th>发布时间</th>
	<th>当前状态</th>
	<th>中标企业</th>
	<th>中标总额</th>
</tr>
<%sql=getsql(request("sql"))
Set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not rs.eof then
	i=0
	do while not rs.eof
	i=i+1%>
<tr align="center">
	<td><%=i%></td>
	<td><%=rs("tendername")%></td>
	<td><%=rs("djh")%></td>
	<td><%=rs("basis")%></td>
	<td><%=rs("tenderclass")%></td>
	<td><%=showtendergrade(rs("tendergradeid"))%></td>
	<td><%=rs("needdate")%></td>
	<td><%response.write rs("startdate")&" "&rs("startdatehour")&"点 至 "&rs("enddate")&" "&rs("enddatehour")&"点"%></td>
	<td><%=shownameint("userrealname","login","id",rs("addman"))%></td>
	<td><%=rs("addtime")%></td>
	<td><%=gettenderstatue(rs("statue"))%></td>
	<td><%finalcomid=rs("finalcompany")
	response.write "<a href='#' onClick='doCheckDetail(""company_main.asp?id="&finalcomid&""","&Modalwidth&","&Modalheight&")' title='点击查看企业详情'>"&shownameint("companyrealname","company","id",finalcomid)&"</a>"%></td>
	<td><%=getfinalmoney(rs("djh"))%></td>
</tr>
	<%Rs.movenext
	loop
else
	response.write "<tr class='grid_span'><td colspan='16'><font color='red'>查无资料纪录！</font></td></tr>"
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