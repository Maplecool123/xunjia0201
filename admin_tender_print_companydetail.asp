<!--#include file="Index_top_print.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
djh=request("djh")
comid=request("comid")
basis=getsql(request("basis"))%>
<table width='100%' cellpadding='0' cellspacing='1' border='0' style=" font-size:10px;word-break: break-all;">
<tr>
	<td align="left" colspan="3">&nbsp;</td>
</tr>
<tr>
	<td align="center" colspan="3"><div onClick="window.print();">询价采购系统供应商报价明细清单</div></td>
</tr>
<tr>
	<td align="center" colspan="3"><%=shownameint("companyrealname","company","id",comid)%></td>
</tr>
<tr>
	<td align="left" width="50%">　系统编号：<%=djh%></td>
	<td align="left" width="20%">&nbsp;</td>
	<td align="left">金额单位：元</td>
</tr>
<tr>
	<td align="left">　采购依据：<%=basis%></td>
	<td align="left"></td>
	<td align="left">打印日期：<%=date()%></td>
</tr>
</table>
<table width='100%' cellpadding='0' cellspacing='1' border='1' style="font-size:10px;word-break: break-all;border-collapse:collapse;">

<tr align="center">
	<td width="5%">序号</td>
	<td width="17%">物资名称</td>
	<td width="17%">规格型号</td>
	<td width="12%">材质</td>
	<td width="8%">单位</td>
	<td width="8%">数量</td>
	<td width="8%">单价</td>
	<td width="12%">到现场价</td>
	<td width="13%">备注</td>
</tr>
<%set rsc=conn.execute("select * from competitive where djh='"&djh&"' and companyid="&comid&" and ifzu=0 order by id")
if not rsc.eof then
	totalshuliang=0
	m=0
	do while not rsc.eof
	m=m+1
	set rs_s=conn.execute("select * from tender where id="&rsc("tenderid")&"")
	if not rs_s.eof then
		response.write "<tr align='center'>"
		response.write "<td>"&m&"</td>"
		response.write "<td align='left'>"&rs_s("material")&"</td>"
		response.write "<td align='center'>"&rs_s("material_guige")&"</td>"
		response.write "<td align='center'>"&rs_s("material_caizhi")&"</td>"
		response.write "<td align='center'>"&rs_s("material_danwei")&"</td>"
		response.write "<td align='center'>"&rs_s("material_shuliang")&"</td>"
		response.write "<td align='center'>"&rsc("singleprice")&"</td>"
		response.write "<td align='center'>"&rsc("spotprice")&"</td>"
		response.write "<td align='center'>&nbsp;</td>"
		response.write "</tr>"
		totalshuliang=totalshuliang+0+rs_s("material_shuliang")
	end if
	rs_s.close
	rsc.movenext
	loop
end if
rsc.close

set rsc=conn.execute("select * from competitive where djh='"&djh&"' and companyid="&comid&" and ifzu=1")
if not rsc.eof then
	response.write "<tr align='center'>"
	response.write "<td>&nbsp;</td>"
	response.write "<td align='left'>合计</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td align='center'>"&totalshuliang&"</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td align='center'>"&rsc("spotprice")&"</td>"
	response.write "<td align='center'>&nbsp;</td>"
	response.write "</tr>"
end if
rsc.close%>
</table>

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