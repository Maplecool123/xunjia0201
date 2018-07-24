<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","竞价信息")

set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
	companyno1=getmaxcompetitivecompanyno(rs_t("djh")," and (statue=0 or statue=1 or statue=2)")
	companyno2=getmaxcompetitivecompanyno(rs_t("djh")," and statue=3")
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<%if rs_t("lastdjh")<>"" then%>
<tr>
	<td class="list_header_required">项目来源</td>
	<td class="list_required"><strong>二次竞价</strong></td>
	<td class="list_header_required">原 项 目</td>
	<td class="list_required"><a href="#" onClick="doCheckDetail('admin_tender_main.asp?id=<%=checkifnum(shownamestrnull("id","tender","djh",rs_t("lastdjh")," and ifzu=1"))%>',<%=Modalwidth%>,<%=Modalheight%>)"><%=rs_t("lastdjh")%></a></td>
</tr>
<%end if%>
<tr>
	<td class="list_header_required" width="10%">项目名称</td>
	<td class="list_required" width="40%"><%=rs_t("tendername")%></td>
	<td class="list_header_required" width="10%">竞价编号</td>
	<td class="list_required" width="40%"><%=rs_t("djh")%></td>
</tr>
<tr>
	<td class="list_header_required">采购依据</td>
	<td class="list_required"><%=rs_t("basis")%></td>
	<td class="list_header_required">项目类型</td>
	<td class="list_required"><%=rs_t("tenderclass")%></td>
</tr>
<tr>
	<td class="list_header_required">竞价等级</td>
	<td class="list_required"><%=showtendergrade(rs_t("tendergradeid"))%></td>
	<td class="list_header_required">需求日期</td>
	<td class="list_required"><%=rs_t("needdate")%></td>
</tr>
<tr>
	<td class="list_header_required">报价开始</td>
	<td class="list_required"><%response.write rs_t("startdate")&" "&rs_t("startdatehour")&"点"%></td>
	<td class="list_header_required">报价截止</td>
	<td class="list_required"><%response.write rs_t("enddate")&" "&rs_t("enddatehour")&"点"%></td>
</tr>
<tr>
	<td class="list_header_required">详情说明</td>
	<td class="list_required" colspan="3"><%=rs_t("detail")%></td>
</tr>
</table>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="20"><strong>采购物资</strong></td>
</tr>
<tr>
	<td class="list_header_required_yellow_center" colspan="2" rowspan="2">物资</td>
	<%redim arr1(companyno1,5)  '可以参选的
	set rsc=conn.execute("select companyid,spotprice,reason,deliverydate,paystyle from competitive where djh='"&rs_t("djh")&"' and ifzu=1 and (statue=0 or statue=1 or statue=2)")
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
	set rsc=conn.execute("select companyid,spotprice,reason,deliverydate,paystyle from competitive where djh='"&rs_t("djh")&"' and ifzu=1 and statue=3")
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
		if checkifnum(rs_t("finalcompany"))-checkifnum(arr1(i,1))=0 then response.write "<font color=red>（中标）</font>"
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
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs_t("djh")&"'")
if not rs_s.eof then
	wzno=0
	do while not rs_s.eof
	wzno=wzno+1%>
<tr>
	<td class="list_header_required_yellow" width="10%">物资<%=wzno%></td>
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
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="4"><strong>竞价条件</strong></td>
</tr>
<tr>
	<td class="list_header_required" width="10%">成功条件</td>
	<td class="list_required" colspan="3">竞价企业数量达到 <strong><%=rs_t("companyno")%></strong> 家&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	竞价最高金额达到 <strong><%=rs_t("minmoney")%></strong> 元</td>
</tr>
</table>
<%end if
rs_t.close%>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->