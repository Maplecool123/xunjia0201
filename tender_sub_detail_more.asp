<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<%if rs("detail")<>"" then%>
<tr class='grid_header'>
	<td><strong>详情说明</strong></td>
</tr>
<tr class="grid_odd">
	<td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("detail")%></td>
</tr>
<%end if%>
</table>
<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<%if rs("finalcompany")<>0 then%>
<tr class='grid_header'>
	<td><strong>中标信息</strong></td>
</tr>
<tr class="grid_odd">
	<td>&nbsp;&nbsp;&nbsp;&nbsp;中标企业：<%=shownameint("companyrealname","company","id",rs("finalcompany"))%>&nbsp;&nbsp;&nbsp;&nbsp;中标时间：<%=rs("decidetime")%>&nbsp;&nbsp;&nbsp;&nbsp;定标人：<%=shownameint("userrealname","login","id",rs("decideman"))%></td>
</tr>
<%end if%>
</table><table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="20"><strong>采购物资（<%=gettotalitem("tender","djh",rs("djh")," and ifzu=0 and ifdel=0")%>项）</strong></td>
</tr>
<tr>
	<td class="list_header_required_yellow_center" colspan="2" rowspan="2">物资</td>
	<%
	companyno1=getmaxcompetitivecompanyno(rs("djh")," and (statue=0 or statue=1 or statue=2)")
	companyno2=getmaxcompetitivecompanyno(rs("djh")," and statue=3")
	redim arr1(companyno1,5)  '可以参选的
	set rsc=conn.execute("select companyid,spotprice,reason,deliverydate,paystyle from competitive where djh='"&rs("djh")&"' and ifzu=1 and (statue=0 or statue=1 or statue=2)")
	if not rsc.eof then
		i=0
		do while not rsc.eof
		i=i+1
		arr1(i,1)=rsc(0)
		arr1(i,2)=rsc(1)
		arr1(i,3)=rsc(2)
		arr1(i,4)=rsc(3)
		arr1(i,5)=rsc(4)
		rsc.movenext
		loop
	end if
	rsc.close
	
	tmp1=""
	tmp2=""
	tmp3=""
	tmp4=""
	tmp5=""
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
				tmp4 = arr1(i,4)
				arr1(i,4) = arr1(j,4)
				arr1(j,4) = tmp4
				tmp5 = arr1(i,5)
				arr1(i,5) = arr1(j,5)
				arr1(j,5) = tmp5
			End If
		Next
	Next
	
	redim arr2(companyno2,5)  '可以参选的
	set rsc=conn.execute("select companyid,spotprice,reason,deliverydate,paystyle from competitive where djh='"&rs("djh")&"' and ifzu=1 and statue=3")
	if not rsc.eof then
		i=0
		do while not rsc.eof
		i=i+1
		arr2(i,1)=rsc(0)
		arr2(i,2)=rsc(1)
		arr2(i,3)=rsc(2)
		arr2(i,4)=rsc(3)
		arr2(i,5)=rsc(4)
		rsc.movenext
		loop
	end if
	rsc.close
	
	tmp1=""
	tmp2=""
	tmp3=""
	tmp4=""
	tmp5=""
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
				tmp4 = arr2(i,4)
				arr2(i,4) = arr2(j,4)
				arr2(j,4) = tmp4
				tmp5 = arr2(i,5)
				arr2(i,5) = arr2(j,5)
				arr2(j,5) = tmp5
			End If
		Next
	Next
		
	For i = 1 To companyno1
		response.write "<td class='list_header_required_center' align='center'><a href='#' onClick='doCheckDetail(""company_main.asp?id="&arr1(i,1)&""","&Modalwidth&","&Modalheight&")' title='点击查看企业详情'>"&shownameint("companyrealname","company","id",arr1(i,1))&"</a>"
		if checkifnum(rs("finalcompany"))-checkifnum(arr1(i,1))=0 then response.write "<font color=red>（中标）</font>"
		response.write "</td>"
	Next
	
	For i = 1 To companyno2
		response.write "<td class='list_header_required_center' align='center'><a href='#' onClick='doCheckDetail(""company_main.asp?id="&arr1(i,1)&""","&Modalwidth&","&Modalheight&")' title='点击查看企业详情'>"&shownameint("companyrealname","company","id",arr2(i,1))&"</a></td>"
	Next%>
</tr>
<tr>
	<%For i = 1 To companyno1
		response.write "<td class='list_header_required_left' align='center'>"
		response.write "付款方式："&arr1(i,5)&"<br>"
		response.write "交货时间："&arr1(i,4)&"<br>"
		response.write "总 金 额："&arr1(i,2)&"元<br>"
		response.write "附加说明："&arr1(i,3)
		response.write "</td>"
	Next
	
	For i = 1 To companyno2
		response.write "<td class='list_header_required_left' align='center'>"
		response.write "付款方式："&arr2(i,5)&"<br>"
		response.write "交货时间："&arr2(i,4)&"<br>"
		response.write "总 金 额："&arr2(i,2)&"元<br>"
		response.write "附加说明："&arr2(i,3)
		response.write "</td>"
	Next%>
</tr>
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs("djh")&"'")
if not rs_s.eof then
	wzno=0
	do while not rs_s.eof
	wzno=wzno+1%>
<tr>
	<td class="list_header_required_yellow" width="5%">物资<%=wzno%></td>
	<td class="list_required">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required" colspan="2">名称：<%=rs_s("material")%></td>
		</tr>
		<tr>
			<td class="list_required" colspan="2">规格：<%=rs_s("material_guige")%></td>
		</tr>
		<tr>
			<td class="list_required" colspan="2">材质：<%=rs_s("material_caizhi")%></td>
		</tr>
		<tr>
			<td class="list_required" width="50%">单位：<%=rs_s("material_danwei")%></td>
			<td class="list_required" width="50%">数量：<%=rs_s("material_shuliang")%></td>
		</tr>
		</table>
	</td>
	
	<%For i = 1 To companyno1
		set rsc=conn.execute("select * from competitive where tenderid="&rs_s("id")&" and companyid="&arr1(i,1)&" and ifzu=0")
		if not rsc.eof then%>
	<td class="list_required">
		<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">

		<tr>
			<td class="list_required">物资单价：<%=sonic(rsc("singleprice"))%> 元</td>
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
	Next
	
	For i = 1 To companyno2
		set rsc=conn.execute("select * from competitive where tenderid="&rs_s("id")&" and companyid="&arr2(i,1)&" and ifzu=0")
		if not rsc.eof then%>
	<td class="list_required">
		<table width='100%' class='list_table' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">

		<tr>
			<td class="list_required">物资单价：<%=sonic(rsc("singleprice"))%> 元</td>
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
	Next%>
</tr>
	<%rs_s.movenext
	loop
end if
rs_s.close%>
</table>