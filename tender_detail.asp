<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
call WhereTable("identity.png","竞价信息")

set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="list_header_required" width="10%">项目名称</td>
	<td class="list_required" width="40%"><%=rs_t("tendername")%>&nbsp;&nbsp;<input type='button' value='打印' onClick="window.print();" class='form_button'></td>
	<td class="list_header_required" width="10%">竞价编号</td>
	<td class="list_required" width="40%"><%=rs_t("djh")%></td>
</tr>
<tr>
	<td class="list_header_required">项目类型</td>
	<td class="list_required"><%=rs_t("tenderclass")%></td>
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
<tr>
	<td class="grid_header" colspan="4"><strong>采购物资（<%=gettotalitem("tender","djh",rs_t("djh")," and ifzu=0 and ifdel=0")%>项）</strong></td>
</tr>
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs_t("djh")&"'")
if not rs_s.eof then
	wzno=0
	do while not rs_s.eof
	wzno=wzno+1
	if (wzno-1) mod 2=0 then response.write "<tr>"%>
	<td class="list_header_required">物资<%=wzno%></td>
	<td class="list_required">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required">名称：<%=rs_s("material")%></td>
		</tr>
		<tr>
			<td class="list_required">规格：<%=rs_s("material_guige")%></td>
		</tr>
		<tr>
			<td class="list_required">材质：<%=rs_s("material_caizhi")%></td>
		</tr>
		<tr>
			<td class="list_required">单位：<%=rs_s("material_danwei")%></td>
		</tr>
		<tr>
			<td class="list_required">数量：<%=rs_s("material_shuliang")%></td>
		</tr>
		</table>
	</td>
	<%if (wzno-1) mod 2=1 then response.write "</tr>"
	rs_s.movenext
	loop
end if
rs_s.close%>
</table>
<%end if
rs_t.close%>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->