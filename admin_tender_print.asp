<!--#include file="Index_top_print.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
set rs=server.createobject("ADODB.RecordSet")
sql="select * from tender where id="&request("id")&""
rs.open sql,conn,1,3
if not rs.eof then
	rs("ifprint")=1
	rs.update
end if
rs.close

set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
	djh=rs_t("djh")
	basis=rs_t("basis")
	companyno1=getmaxcompetitivecompanyno(rs_t("djh")," and (statue=0 or statue=1 or statue=2)")
	companyno2=getmaxcompetitivecompanyno(rs_t("djh")," and statue=3")%>
<table width='100%' cellpadding='0' cellspacing='1' border='0' style=" font-size:10px;word-break: break-all;">
<tr>
	<td align="center" colspan="3"><div onClick="window.print();"><font style="font-size:24px;">询价采购系统报价排名情况</font></div></td>
</tr>
<tr>
	<td align="left" width="50%">　系统编号：<%=rs_t("djh")%></td>
	<td align="left" width="20%">&nbsp;</td>
	<td align="left">金额单位：元</td>
</tr>
<tr>
	<td align="left" colspan="3">　项目类型：<%=rs_t("tenderclass")%></td>
</tr>
<tr>
	<td align="left">　采购依据：<%=rs_t("basis")%></td>
	<td align="left"></td>
	<td align="left">打印日期：<%=date()%></td>
</tr>
</table>
<table width='100%' cellpadding='0' cellspacing='1' border='1' style="font-size:10px;word-break: break-all;border-collapse:collapse;">
<%redim arr1(companyno1,3)  '可以参选的
set rsc=conn.execute("select companyid,spotprice,reason from competitive where djh='"&rs_t("djh")&"' and ifzu=1 and (statue=0 or statue=1 or statue=2)")
if not rsc.eof then
	i=0
	do while not rsc.eof
	i=i+1
	arr1(i,1)=rsc(0)
	arr1(i,2)=rsc(1)
	arr1(i,3)=rsc(2)
	rsc.movenext
	loop
end if
rsc.close

tmp1=""
tmp2=""
tmp3=""
For i = 1 To companyno1  '从小到大排列
	For j = i + 1 To companyno1
		If arr1(i,2)-arr1(j,2)>0 Then
			tmp1 = arr1(i,2)
			arr1(i,2) = arr1(j,2)
			arr1(j,2) = tmp1
			tmp2 = arr1(i,1)
			arr1(i,1) = arr1(j,1)
			arr1(j,1) = tmp2
			tmp3 = arr1(i,3)
			arr1(i,3) = arr1(j,3)
			arr1(j,3) = tmp3
		End If
	Next
Next

redim arr2(companyno2,3)  '可以参选的
set rsc=conn.execute("select companyid,spotprice,reason from competitive where djh='"&rs_t("djh")&"' and ifzu=1 and statue=3")
if not rsc.eof then
	i=0
	do while not rsc.eof
	i=i+1
	arr2(i,1)=rsc(0)
	arr2(i,2)=rsc(1)
	arr2(i,3)=rsc(2)
	rsc.movenext
	loop
end if
rsc.close

tmp1=""
tmp2=""
tmp3=""
For i = 1 To companyno2  '从小到大排列
	For j = i + 1 To companyno2
		If arr2(i,2)-arr2(j,2)>0 Then
			tmp1 = arr2(i,2)
			arr2(i,2) = arr2(j,2)
			arr2(j,2) = tmp1
			tmp2 = arr2(i,1)
			arr2(i,1) = arr2(j,1)
			arr2(j,1) = tmp2
			tmp3 = arr2(i,3)
			arr2(i,3) = arr2(j,3)
			arr2(j,3) = tmp3
		End If
	Next
Next
%>
<tr align="center">
	<td width="5%">序号</td>
	<td width="50%">供应商</td>
	<td width="10%">竞价价格</td>
	<td width="10%">排名情况</td>
	<td width="25%">备注</td>
</tr>
<%no=0
For i = 1 To companyno1
	no=no+1
	response.write "<tr align='center'>"
	response.write "<td>"&no&"</td>"
	response.write "<td align='left' onClick='doCheckDetail(""admin_tender_print_companydetail.asp?comid="&arr1(i,1)&"&djh="&rs_t("djh")&"&basis="&Filtersql(rs_t("basis"))&""","&Modalheight&","&Modalwidth&")'>"&shownameint("companyrealname","company","id",arr1(i,1))&"</td>"
	response.write "<td align='center'>"&arr1(i,2)&"</td>"
	response.write "<td align='center'>"&no&"</td>"
	response.write "<td align='center'>"&arr1(i,3)&"</td>"
	response.write "</tr>"
Next

For i = 1 To companyno2
	no=no+1
	response.write "<tr align='center'>"
	response.write "<td>"&no&"</td>"
	response.write "<td align='left' onClick='doCheckDetail(""admin_tender_print_companydetail.asp?comid="&arr2(i,1)&"&djh="&rs_t("djh")&"&basis="&Filtersql(rs_t("basis"))&""","&Modalheight&","&Modalwidth&")'>"&shownameint("companyrealname","company","id",arr2(i,1))&"（不达标）</td>"
	response.write "<td align='center'>"&arr2(i,2)&"</td>"
	response.write "<td align='center'><font color='#FF0000'>"&no&"</font></td>"
	response.write "<td align='center'>"&arr2(i,3)&"</td>"
	response.write "</tr>"
Next%>
</table>
<%end if
rs_t.close%>
<table width='100%' cellpadding='0' cellspacing='1' border='0' style=" font-size:10px;word-break: break-all;">
<tr>
	<td align="right">本数据出自<%=sysname%></td>
</tr>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->