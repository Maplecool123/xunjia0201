<!--#include file="Index_top.asp" -->
<%
call WhereTable("identity.png","企业详细信息")

set rsm=conn.execute("select * from company where id="&request("id")&"")
if not rsm.eof then%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all">
<tr>
	<td class="list_header_required">企业名称</td>
	<td class="list_required" colspan="3"><%=rsm("companyrealname")%></td>
	<td class="list_header_required">登陆用户名</td>
	<td class="list_required"><%=rsm("companyname")%></td>
</tr>
<tr>
	<td class="list_header_required" width="13%">审核状态</td>
	<td class="list_required" width="20%"><%=getcompanystatue(rsm("statue"))%></td>
	<td class="list_header_required" width="13%">企业类型</td>
	<td class="list_required" width="21%"><%=shownameint("companytypename","companytype","id",rsm("companytype"))%></td>
	<td class="list_header_required" width="13%">机构代码证号</td>
	<td class="list_required" width="21%"><%=rsm("companyorgno")%></td>
</tr>
<tr>
	<td class="list_header_required">税号</td>
	<td class="list_required"><%=rsm("companytaxno")%></td>
	<td class="list_header_required">法定代表人</td>
	<td class="list_required"><%=rsm("companymanager")%></td>
	<td class="list_header_required">联系人姓名</td>
	<td class="list_required"><%=rsm("companycontactman")%></td>
</tr>
<tr>
	<td class="list_header_required">联系电话</td>
	<td class="list_required"><%=rsm("companycontacttel")%></td>
	<td class="list_header_required">公司地址</td>
	<td class="list_required" colspan="3"><%=rsm("companyaddress")%></td>
</tr>
<tr>
	<td class="list_header_required">电子邮件</td>
	<td class="list_required"><%=rsm("companyemail")%></td>
	<td class="list_header_required">主营业务</td>
	<td class="list_required" colspan="3"><%=rsm("companymajor")%></td>
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
</table>

<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all">
<tr>
	<td class="list_header_required" width="13%">竞价记录信息</td>
	<td class="list_required">
	<%call getcompanycompetitiveno(request("id"))   '竞价信息
	dim competitiveno,getcompetitiveno,getcompetitiveallprice,gettime%>
	参加竞价次数：<%=competitiveno%>次
	中标次数：<%=getcompetitiveno%>次
	中标总额：<%=getcompetitiveallprice%>元
	上次中标：<%=gettime%>
	</td>
</tr>
</table>
<%end if
rsm.close%>
<!--#include file="Index_bottom.asp" -->