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
%>
<html> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<title>物资汇总表</title>
</head>
<body>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
	<th>序号</th>
	<th>采购依据</th>
	<th>项目类型</th>
	<th>物资名称</th>
	<th>物资规格</th>
	<th>物资材质</th>
	<th>单位</th>
	<th>数量</th>
	<th>单价</th>
	<th>到现场价</th>
	<th>中标人供应商</th>
	<th>计划员</th>
	<th>中标日期</th>
	<th>备注</th>
</tr>
<%sql=getsql(request("sql"))
Set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not rs.eof then
	i=0
	do while not rs.eof
	set rsc=conn.execute("select * from tender where djh='"&rs("djh")&"' and ifzu=0")
	if not rsc.eof then
		do while not rsc.eof
		i=i+1
		
		set rsf=conn.execute("select * from competitive where tenderid="&rsc("id")&" and (statue=1 or statue=2) and ifzu=0")%>
<tr align="center">
	<td><%=i%></td>
	<td><%=rs("basis")%></td>
	<td><%=rs("tenderclass")%></td>
	<td><%=rsc("material")%></td>
	<td><%=rsc("material_guige")%></td>
	<td><%=rsc("material_caizhi")%></td>
	<td><%=rsc("material_danwei")%></td>
	<td><%=rsc("material_shuliang")%></td>
	<td><%=rsf("singleprice")%></td>
	<td><%=rsf("spotprice")%></td>
	<td><%=shownameint("companyrealname","company","id",rsf("companyid"))%></td>
	<td><%=shownameint("userrealname","login","id",rs("addman"))%></td>
	<td><%=rsf("gettime")%></td>
	<td>&nbsp;</td>
</tr>
		<%rsf.close
		
		rsc.movenext
		loop
	end if
	rsc.close
	
	Rs.movenext
	loop
else
	response.write "<tr><td colspan='14'><font color='red'>查无资料纪录！</font></td></tr>"
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