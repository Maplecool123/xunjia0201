<!--#include file="Index_top.asp" -->
<!--#include file="inc/Md5.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-2=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","审核基本信息")

gourl="company_class="&request.form("company_class")&"&company_type="&request.form("company_type")&"&checkstatue="&request.form("checkstatue")&"&checkcondition="&request.form("checkcondition")&"&checkstyle="&request.form("checkstyle")&"&keyword1="&request.form("keyword1")&"&page="&request.form("page")

if request.form("action")="saveedit" then
	nowstatue=checkifnum(request.form("statue"))
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from company where id="&request.form("id")&""
	rs.open sql,conn,1,3
	if checkifnum(rs("ifshen1"))=0 then
		rs("ifshen1")=nowstatue
		rs("shenman1")=session("user_id")
		rs("shentime1")=now()
		
		if nowstatue-2=0 then
			rs("statue")=2
		else
			rs("statue")=0
		end if
	elseif checkifnum(rs("ifshen1"))>0 then
		if nowstatue>0 then
			rs("ifshen2")=nowstatue
			rs("shenman2")=session("user_id")
			rs("shentime2")=now()
			
			if nowstatue-1=0 and rs("ifshen1")-1=0 then
				rs("statue")=1
			else
				rs("statue")=2
			end if
		end if
	end if
	rs.update
	rs.close
	
	call HintAndTurn("审核成功！","admin_company_shen.asp?"&gourl)
end if

set rsm=conn.execute("select * from company where id="&request("id")&"")
if not rsm.eof then%>
<form name='form' method="post">
<input type="hidden" name="action" value="saveedit">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="company_class" value="<%=request("company_class")%>">
<input type="hidden" name="company_type" value="<%=request("company_type")%>">
<input type="hidden" name="checkstatue" value="<%=request("checkstatue")%>">
<input type="hidden" name="checkcondition" value="<%=request("checkcondition")%>">
<input type="hidden" name="checkstyle" value="<%=request("checkstyle")%>">
<input type="hidden" name="keyword1" value="<%=request("keyword1")%>">
<input type="hidden" name="page" value="<%=request("page")%>">
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<tr>
	<td class="list_header_required">审核结果</td>
	<td class="list_required" colspan="5">
	<%if rsm("ifshen1")-2=0 then
		response.write "第一次审核未通过"
	elseif rsm("ifshen1")-1=0 then
		response.write "已通过一次审核"
		if checkifnum(rsm("shenman1"))-checkifnum(session("user_id"))=0 then
			response.write "<br>您已经审核过一次了，不可以重复审核！"
		else
			response.write "<br>通过<input type='radio' name='statue' class='form_radio_normal' value='1' checked>"
			response.write "不通过<input type='radio' name='statue' class='form_radio_normal' value='2'>"
		end if
	else
		response.write "<br>通过<input type='radio' name='statue' class='form_radio_normal' value='1' checked>"
		response.write "不通过<input type='radio' name='statue' class='form_radio_normal' value='2'>"
	end if%>
	<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">企业名称</td>
	<td class="list_required" colspan="5"><%=rsm("companyrealname")%></td>
</tr>
<tr>
	<td class="list_header_required" width="13%">登陆用户名</td>
	<td class="list_required" width="20%"><%=rsm("companyname")%></td>
	<td class="list_header_required" width="13%">企业归类</td>
	<td class="list_required" width="20%"><%=rsm("companyclass")%></td>
	<td class="list_header_required" width="13%">企业类型</td>
	<td class="list_required" width="21%"><%=shownameint("companytypename","companytype","id",rsm("companytype"))%></td>
</tr>
<tr>
	<td class="list_header_required">机构代码证号</td>
	<td class="list_required"><%=rsm("companyorgno")%></td>
	<td class="list_header_required">税号</td>
	<td class="list_required"><%=rsm("companytaxno")%></td>
	<td class="list_header_required">企业级别</td>
	<td class="list_required"><%=shownameint("companygradename","companygrade","id",rsm("companygrade"))%></td>
</tr>
<tr>
	<td class="list_header_required">法定代表人</td>
	<td class="list_required"><%=rsm("companymanager")%></td>
	<td class="list_header_required">联系人姓名</td>
	<td class="list_required"><%=rsm("companycontactman")%></td>
	<td class="list_header_required">联系电话</td>
	<td class="list_required"><%=rsm("companycontacttel")%></td>
</tr>
<tr>
	<td class="list_header_required">公司地址</td>
	<td class="list_required" colspan="3"><%=rsm("companyaddress")%></td>
	<td class="list_header_required">电子邮件</td>
	<td class="list_required"><%=rsm("companyemail")%></td>
</tr>
<tr>
	<td class="list_header_required">主营业务</td>
	<td class="list_required" colspan="5"><%=rsm("companymajor")%></td>
</tr>
<tr>
	<td class="list_header_required">备注</td>
	<td class="list_required" colspan="5"><%=rsm("companybeizhu")%></td>
</tr>
</table>

<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all">
<tr>
	<td class="list_header_required" width="13%">企业资质盖章图片</td>
	<td class="list_required">
		<%companyfile=rsm("companyfile")
		if companyfile<>"" then
			companyfiles=split(companyfile,",")
			response.write "<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style='word-break: break-all;'>"
			response.write "<tr align='center'>"
			response.write "<td class='list_required' width='25%'>竞价采购承诺书</td>"
			response.write "<td class='list_required' width='25%'>企业法人营业执照</td>"
			response.write "<td class='list_required' width='25%'>税务登记证</td>"
			response.write "<td class='list_required' width='25%'>组织机构代码证</td>"
			response.write "</tr>"
			
			response.write "<tr align='center'>"
			for i=0 to ubound(companyfiles)
				response.write "<td class='list_required'><img src='uploadfile/companyfile/"&companyfiles(i)&"' width='120' height='160' border=0 title='点击看大图' onClick='doCheckDetailEx(""uploadfile/companyfile/"&companyfiles(i)&""","&Modalwidth&","&Modalheight&")'></td>"
				if i=3 then exit for
			next
			response.write "</tr>"
			for i=4 to ubound(companyfiles)
				if i mod 4=0 then response.write "<tr align='center'>"
				response.write "<td class='list_required'><img src='uploadfile/companyfile/"&companyfiles(i)&"' width='120' height='160' border=0 title='点击看大图' onClick='doCheckDetailEx(""uploadfile/companyfile/"&companyfiles(i)&""","&Modalwidth&","&Modalheight&")'></td>"
				if i mod 4=3 then response.write "</tr>"
			next
			response.write "</table>"
		else
			response.write "暂未上传盖章资质扫描图片"
		end if%>
	</td>
</tr>
</table>

<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all">
<tr>
	<td class="list_header_required" width="13%">企业合同盖章图片</td>
	<td class="list_required">
		<%companycontract=rsm("companycontract")
		if companycontract<>"" then
			companycontracts=split(companycontract,",")
			response.write "<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style='word-break: break-all;'>"
			for i=0 to ubound(companycontracts)
				if i mod 4=0 then response.write "<tr align='center'>"
				response.write "<td width='20%' class='list_required'><img src='uploadfile/companycontract/"&companycontracts(i)&"' width='120' height='160' border=0 title='点击看大图' onClick='doCheckDetailEx(""uploadfile/companycontract/"&companycontracts(i)&""","&Modalwidth&","&Modalheight&")'></td>"
				if i mod 4=3 then response.write "</tr>"
			next
			response.write "</table>"
		else
			response.write "暂未上传盖章资质扫描图片"
		end if%>
	</td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 确认 '>
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</table>
</form>
<%end if
rsm.close
else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->