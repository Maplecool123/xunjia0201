<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
call WhereTable("identity.png","报价信息")

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
</table>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="4"><strong>报价相关信息</strong></td>
</tr>
<%deliverydate=""
paystyle=""
beizhu=""
alltotalprice=0
allspotprice=0
set rs_c=conn.execute("select deliverydate,paystyle,beizhu,totalprice,spotprice from competitive where djh='"&rs_t("djh")&"' and companyid="&session("user_id")&" and ifzu=1")
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
<tr>
	<td class="grid_header" colspan="4"><strong>采购物资报价</strong></td>
</tr>
<tr>
	<td class="list_header_required">物资总价<br>合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;计</td>
	<td class="list_required"><%=alltotalprice%>元</td>
	<td class="list_header_required">到现场价<br>合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;计</td>
	<td class="list_required"><%=allspotprice%>元</td>
</tr>
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs_t("djh")&"'")
if not rs_s.eof then
	wzno=0
	do while not rs_s.eof
	wzno=wzno+1%>
<tr>
	<td class="list_header_required" width="10%">物资<%=wzno%></td>
	<td class="list_required" width="40%">
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
	<%singleprice=0
	totalprice=0
	spotprice=0
	set rs_c=conn.execute("select * from competitive where tenderid="&rs_s("id")&" and companyid="&session("user_id")&" and ifzu=0")
	if not rs_c.eof then
		singleprice=rs_c("singleprice")
		totalprice=rs_c("totalprice")
		spotprice=rs_c("spotprice")
	end if
	rs_c.close%>
	<td class="list_header_required" width="10%">物资报价</td>
	<td class="list_required">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required">物资单价：<%=singleprice%>元</td>
		</tr>
		<tr>
			<td class="list_required">物资总价：<%=totalprice%>元</td>
		</tr>
		<tr>
			<td class="list_required">到现场价：<%=spotprice%>元</td>
		</tr>
		</table>
	</td>
</tr>
	<%rs_s.movenext
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