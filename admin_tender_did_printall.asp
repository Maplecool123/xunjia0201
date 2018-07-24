<!--#include file="Index_top_print.asp" -->
<% 
if session("user_id")="" then
	response.write "<script language='javascript'>alert('您还没有登陆，请先登陆');"
	response.write "window.location ='ad_login.asp';</script> "
	Response.End 
end if
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
<title>物资汇总表 - 打印</title>
<table width='100%' cellpadding='0' cellspacing='1' border='0' style=" font-size:10px;word-break: break-all;">
<tr>
	<td align="center" colspan="14" height="30"><div onClick="window.print();">物资汇总表</div></td>
</tr>
</table>
<table width='100%' cellpadding='0' cellspacing='1' border='1' style="font-size:10px;word-break: break-all;border-collapse:collapse;">
<tr align="center">
	<td>序号</td>
	<td>采购依据</td>
	<td>项目类型</td>
	<td>物资名称</td>
	<td>物资规格</td>
	<td>物资材质</td>
	<td>单位</td>
	<td>数量</td>
	<td>单价</td>
	<td>到现场价</td>
	<td>中标人供应商</td>
	<td>计划员</td>
	<td>中标日期</td>
	<td>备注</td>
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
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->