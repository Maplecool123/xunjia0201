<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","竞价信息")

set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;"><%if rs_t("lastdjh")<>"" then%>
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
<tr>
	<td class="grid_header" colspan="4"><strong>采购物资</strong></td>
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
<tr>
	<td class="grid_header" colspan="4"><strong>竞价条件</strong></td>
</tr>
<tr>
	<td class="list_header_required">成功条件</td>
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