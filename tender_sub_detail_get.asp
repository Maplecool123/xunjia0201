<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<%if rs("detail")<>"" then%>
<tr class='grid_header'>
	<td><strong>详情说明</strong></td>
</tr>
<tr class="grid_odd">
	<td><%=rs("detail")%></td>
</tr>
<%end if%>
</table>
<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="4"><strong>报价相关信息</strong></td>
</tr>
<%deliverydate=""
paystyle=""
beizhu=""
alltotalprice=0
allspotprice=0
set rs_c=conn.execute("select deliverydate,paystyle,beizhu,totalprice,spotprice from competitive where djh='"&rs("djh")&"' and companyid="&session("user_id")&" and ifzu=1")
if not rs_c.eof then
	deliverydate=rs_c(0)
	paystyle=rs_c(1)
	beizhu=rs_c(2)
	alltotalprice=checkifnum(rs_c(3))
	allspotprice=checkifnum(rs_c(4))
end if
rs_c.close%>
<tr>
	<td class="list_header_required">交货时间</td>
	<td class="list_required"><%=deliverydate%></td>
	<td class="list_header_required">付款方式</td>
	<td class="list_required"><%=paystyle%></td>
</tr>
<tr>
	<td class="list_header_required">备注</td>
	<td class="list_required" colspan="3"><%=beizhu%></td>
</tr>
<tr class='grid_header'>
	<td colspan="4"><strong>竞价物资</strong></td>
</tr>
<tr>
	<td class="list_header_required" width="10%">物资总价<br>合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;计</td>
	<td class="list_required" width="40%"><%=alltotalprice%>元</td>
	<td class="list_header_required" width="10%">到现场价<br>合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;计</td>
	<td class="list_required" width="40%"><%=allspotprice%>元</td>
</tr>
<%set rs_sub=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs("djh")&"'")
if not rs_sub.eof then
	wzno=0
	do while not rs_sub.eof
	wzno=wzno+1
	response.write "<tr>"%>
	<td class="list_header_required_yellow" width="10%">物资 <%=wzno%></td>
	<td class="list_required" width="40%">
		<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
		<tr>
			<td class="list_required" width="100%" colspan="2">名称：<%=rs_sub("material")%></td>
		</tr>
		<tr>
			<td class="list_required" colspan="2">规格：<%=rs_sub("material_guige")%></td>
		</tr>
		<tr>
			<td class="list_required" colspan="2">材质：<%=rs_sub("material_caizhi")%></td>
		</tr>
		<tr>
			<td class="list_required" width="50%">单位：<%=rs_sub("material_danwei")%></td>
			<td class="list_required" width="50%">数量：<%=rs_sub("material_shuliang")%></td>
		</tr>
		</table>
	</td>
	<%set rsc=conn.execute("select * from competitive where tenderid="&rs_sub("id")&" and companyid="&session("user_id")&"")
	if not rsc.eof then%>
	<td class="list_required" colspan="2">
		<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
		<tr>
			<td class="list_required" width="100%">物资单价：<%=sonic(rsc("singleprice"))%> 元</td>
		</tr>
		<tr>
			<td class="list_required">物资总价：<%=sonic(rsc("totalprice"))%> 元</td>
		</tr>
		<tr>
			<td class="list_required">到现场价：<%=sonic(rsc("spotprice"))%> 元</td>
		</tr>
		</table>
	</td>
	<%end if
	rsc.close
	
	response.write "</tr>"
	rs_sub.movenext
	loop
end if
rs_sub.close%>
</table>