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
<tr class='grid_header'>
	<td colspan="4"><strong>竞价物资（<%=gettotalitem("tender","djh",rs("djh")," and ifzu=0 and ifdel=0")%>项）</strong></td>
</tr>
<%set rs_sub=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs("djh")&"'")
if not rs_sub.eof then
	wzno=0
	do while not rs_sub.eof
	wzno=wzno+1
	if (wzno-1) mod 2=0 then response.write "<tr>"%>
	<td class="list_header_required_yellow" width="10%">物资 <%=wzno%></td>
	<td class="list_required" width="40%">
		<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
		<tr>
			<td class="list_required" width="100%"><%=wzno%>名称：<%=rs_sub("material")%></td>
		</tr>
		<tr>
			<td class="list_required">规格：<%=rs_sub("material_guige")%></td>
		</tr>
		<tr>
			<td class="list_required">材质：<%=rs_sub("material_caizhi")%></td>
		</tr>
		<tr>
			<td class="list_required">单位：<%=rs_sub("material_danwei")%></td>
		</tr>
		<tr>
			<td class="list_required">数量：<%=rs_sub("material_shuliang")%></td>
		</tr>
		</table>
	</td>
	<%if (wzno-1) mod 2=1 then response.write "</tr>"
	rs_sub.movenext
	loop
end if
rs_sub.close%>
</table>