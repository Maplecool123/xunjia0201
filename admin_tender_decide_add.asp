<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","定标")

gourl="needtime1="&request("needtime1")&"&needtime2="&request("needtime2")&"&needtime3="&request("needtime3")&"&needtime4="&request("needtime4")&"&start_date="&request("start_date")&"&end_date="&request("end_date")&"&tender_class="&request("tender_class")&"&checkgrade="&request("checkgrade")&"&checkstyle="&request("checkstyle")&"&checkcondition="&request("checkcondition")&"&checkstatue="&request("checkstatue")&"&keyword1="&request("keyword1")&"&page="&request("page")

nowfinalcompanyid=checkifnum(request("finalcompanyid"))
nowcomid=replace(request("comid")," ","")
if nowcomid<>"" then comids=split(nowcomid,",")
	
if request("action")="saveadd" then
	if request("ifliubiao")="2" then
		if request("finalcompanyid")="" then
			call HintAndBack("请选择中标企业！",1)
		end if
		
		for i=0 to ubound(comids)
			if request("notgood"&comids(i))="yes" and request("reason"&comids(i))="" then
				call HintAndBack("不达标企业必须填写备注说明！",1)
				exit for
			end if
		next
			
		set rs=server.createobject("ADODB.RecordSet")
		sql="select * from [tender] where djh='"&request("djh")&"'"
		rs.open sql,conn,1,3
		if not rs.eof then
			do while not rs.eof
			rs("finalcompany")=nowfinalcompanyid
			rs("decideman")=session("user_id")
			rs("decidetime")=now()
			rs("statue")=2
			rs.update
			rs.movenext
			loop
		end if
		rs.close
		
		for i=0 to ubound(comids)
			set rs=server.createobject("ADODB.RecordSet")
			sql="select * from [competitive] where djh='"&request("djh")&"' and companyid="&comids(i)&""
			rs.open sql,conn,1,3
			if not rs.eof then
				do while not rs.eof
				if request("notgood"&comids(i))="yes" then
					rs("statue")=3
					rs.update
				end if
				rs.movenext
				loop
			end if
			rs.close
			
			set rs=server.createobject("ADODB.RecordSet")
			sql="select * from [competitive] where djh='"&request("djh")&"' and companyid="&comids(i)&" and ifzu=1"
			rs.open sql,conn,1,3
			if not rs.eof then
				rs("reason")=request("reason"&comids(i))
				rs.update
			end if
			rs.close
		next
		
		set rs=server.createobject("ADODB.RecordSet")
		sql="select * from [competitive] where djh='"&request("djh")&"' and companyid="&nowfinalcompanyid&""
		rs.open sql,conn,1,3
		if not rs.eof then
			do while not rs.eof
			rs("gettime")=now()
			rs("statue")=1
			rs.update
			rs.movenext
			loop
		end if
		rs.close
				
		call HintAndTurn("恭喜，定标完成！",request("fromurl")&".asp?"&gourl)
	else
		set rs=server.createobject("ADODB.RecordSet")
		sql="select * from [tender] where djh='"&request("djh")&"'"
		rs.open sql,conn,1,3
		if not rs.eof then
			do while not rs.eof
			rs("finalcompany")=0
			rs("decideman")=session("user_id")
			rs("decidetime")=now()
			rs("statue")=3
			rs.update
			rs.movenext
			loop
		end if
		rs.close
		
		call HintAndTurn("此次采购竞价流标了！",request("fromurl")&".asp?"&gourl)
	end if
end if

set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
	djh=rs_t("djh")
	basis=rs_t("basis")
	companyno1=getmaxcompetitivecompanyno(rs_t("djh"),"")%>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<input type="hidden" name="id" value="<%=rs_t("id")%>">
<input type="hidden" name="djh" value="<%=rs_t("djh")%>">
<input type="hidden" name="fromurl" value="<%=request("fromurl")%>">
<input type="hidden" name="needtime1" value="<%=request("needtime1")%>">
<input type="hidden" name="needtime2" value="<%=request("needtime2")%>">
<input type="hidden" name="needtime3" value="<%=request("needtime3")%>">
<input type="hidden" name="needtime4" value="<%=request("needtime4")%>">
<input type="hidden" name="start_date" value="<%=request("start_date")%>">
<input type="hidden" name="end_date" value="<%=request("end_date")%>">
<input type="hidden" name="company_class" value="<%=request("company_class")%>">
<input type="hidden" name="company_type" value="<%=request("company_type")%>">
<input type="hidden" name="checkgrade" value="<%=request("checkgrade")%>">
<input type="hidden" name="checkstatue" value="<%=request("checkstatue")%>">
<input type="hidden" name="checkcondition" value="<%=request("checkcondition")%>">
<input type="hidden" name="checkstyle" value="<%=request("checkstyle")%>">
<input type="hidden" name="keyword1" value="<%=request("keyword1")%>">
<input type="hidden" name="page" value="<%=request("page")%>">
<table width='100%' cellpadding='0' cellspacing='1' border='0' style=" font-size:10px;word-break: break-all;">
<tr>
	<td align="center" colspan="3" class="grid_header"><strong>询价采购系统报价定标</strong></td>
</tr><tr align="center">
	<td class="list_header_required_left" colspan="3">是否流标：否<input type='radio' class='form_radio_normal' name='ifliubiao' id='ifliubiao' value='2' checked> 是<input type='radio' class='form_radio_normal' name='ifliubiao' id='ifliubiao' value='3' onClick="alert('您选择了流标，请注意！');"></td>
</tr>
<tr>
	<td align="left" width="50%" class="list_header_required_left">系统编号：<%=rs_t("djh")%></td>
	<td align="left" width="20%" class="list_header_required_left">&nbsp;</td>
	<td align="left" class="list_header_required_left">金额单位：元</td>
</tr>
<tr>
	<td align="left" colspan="3" class="list_header_required_left">项目类型：<%=rs_t("tenderclass")%></td>
</tr>
<tr>
	<td align="left" class="list_header_required_left">采购依据：<%=rs_t("basis")%></td>
	<td align="left" class="list_header_required_left">&nbsp;</td>
	<td align="left" class="list_header_required_left">打印日期：<%=date()%></td>
</tr>
</table>
<table width='100%' cellpadding='0' cellspacing='1' border='1' style="font-size:10px;word-break: break-all;border-collapse:collapse;">
<%redim arr1(companyno1,3)  '可以参选的
set rsc=conn.execute("select companyid,spotprice,reason from competitive where djh='"&rs_t("djh")&"' and ifzu=1")
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
%>
<tr align="center">
	<td width="5%" class="list_header_required_yellow_center">序号</td>
	<td width="40%" class="list_header_required_yellow_center">供应商</td>
	<td width="10%" class="list_header_required_yellow_center">竞价价格</td>
	<td width="10%" class="list_header_required_yellow_center">排名情况</td>
	<td width="22%" class="list_header_required_yellow_center">备注</td>
	<td width="5%" class="list_header_required_yellow_center">定标</td>
	<td width="8%" class="list_header_required_yellow_center">不达标</td>
</tr>
<%no=0
For i = 1 To companyno1
	no=no+1
	response.write "<tr align='center'>"
	response.write "<td>"&no&"</td>"
	response.write "<td align='left' onClick='doCheckDetail(""admin_tender_print_companydetail.asp?comid="&arr1(i,1)&"&djh="&rs_t("djh")&"&basis="&rs_t("basis")&""","&Modalheight&","&Modalwidth&")'>"&shownameint("companyrealname","company","id",arr1(i,1))&"</td>"
	response.write "<td align='center'>"&arr1(i,2)&"</td>"
	response.write "<td align='center'>"&no&"</td>"
	response.write "<td align='center'><input type='text' class='form_textbox_normal' name='reason"&arr1(i,1)&"' id='reason"&arr1(i,1)&"' value='' style='width:200px;'></td>"
	response.write "<td align='center'><input type='radio' class='form_radio_normal' name='finalcompanyid' id='finalcompanyid' value='"&arr1(i,1)&"'></td>"
	response.write "<td align='center'><input type='checkbox' class='form_checkbox_normal' name='notgood"&arr1(i,1)&"' id='notgood"&arr1(i,1)&"' value='yes'></td>"
	response.write "</tr>"
	response.write "<input type='hidden' name='comid' id='comid' value='"&arr1(i,1)&"'>"
Next%>
</table>
<table width='100%' cellpadding='0' cellspacing='1' border='1' style="font-size:10px;word-break: break-all;border-collapse:collapse;">
<tr>
	<td class="list_header_command" width="25%">指令</td>
	<td class="list_command"><input class='form_submit' type='submit' name='cmd' value=' 确定 '>
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</table>
</form>
<%end if
rs_t.close%>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->