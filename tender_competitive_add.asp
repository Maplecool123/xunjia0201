<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
call WhereTable("identity.png","竞价信息")

gourl="needtime1="&request.form("needtime1")&"&needtime2="&request.form("needtime2")&"&needtime3="&request.form("needtime3")&"&needtime4="&request.form("needtime4")&"&start_date="&request.form("start_date")&"&end_date="&request.form("end_date")&"&tender_class="&request.form("tender_class")&"&checkgrade="&request.form("checkgrade")&"&checkstyle="&request.form("checkstyle")&"&checkcondition="&request.form("checkcondition")&"&checkstatue="&request.form("checkstatue")&"&keyword1="&request.form("keyword1")&"&page="&request.form("page")

if request.form("action")="saveadd" then
	nowdjh=request.form("djh")
	nowtenderid=replace(request.form("tenderid")," ","")
	
	set rs_t=conn.execute("select * from tender where id="&request.form("id")&"")
	if not rs_t.eof then
		if (rs_t("startdate")=date() and rs_t("startdatehour")>hour(now()) or rs_t("startdate")>date()) then
			call HintAndBack("很抱歉，尚未到竞价开始时间！",1)
		elseif (rs_t("enddate")=date() and rs_t("enddatehour")<=hour(now()) or rs_t("enddate")<date()) then
			call HintAndBack("很抱歉，竞价截止时间已过！",1)
		end if
	else
		call HintAndBack("很抱歉，该竞价项目可能已经被取消了！",1)
	end if
	rs_t.close
	
	nowcompanyid=request.form("companyid")
	nowdeliverydate=request.form("deliverydate")
	nowpaystyle=request.form("paystyle")
	nowbeizhu=request.form("beizhu")
			
	if nowtenderid<>"" then
		alltotalprice=0
		allspotprice=0
		tenderids=split(nowtenderid,",")
		
		set rs=server.createobject("ADODB.RecordSet")
		sql="select * from competitive where djh='"&nowdjh&"' and companyid="&nowcompanyid&""
		rs.open sql,conn,1,3
		if not rs.eof then
			call HintAndBack("您已经参加过本项目的竞价，请核实！",1)
			response.end
		end if
		rs.close
			
		for i=0 to ubound(tenderids)
			nowsingleprice=request.form("singleprice"&tenderids(i))
			nowtotalprice=request.form("totalprice"&tenderids(i))
			nowspotprice=request.form("spotprice"&tenderids(i))
			
			if nowsingleprice="" or nowspotprice="" then
				call HintAndBack("每项物资都必须报价，且所有内容均需要填写，请重新输入！",1)
			end if
		next
		
		for i=0 to ubound(tenderids)
			nowsingleprice=checkifnum(request.form("singleprice"&tenderids(i)))
			nowtotalprice=checkifnum(request.form("totalprice"&tenderids(i)))
			nowspotprice=checkifnum(request.form("spotprice"&tenderids(i)))
			
			set rsf=conn.execute("select material_shuliang from tender where id="&tenderids(i))
			if not rsf.eof then
				material_shuliang=checkifnum(rsf(0))
			end if
			rsf.close
			
			counttotalprice=sishewuru(material_shuliang*nowsingleprice,2)
			countspotprice=sishewuru(material_shuliang*nowsingleprice,2)
			
			if abs(counttotalprice-nowtotalprice)>=1 then
				nowtotalprice=counttotalprice
				nowspotprice=countspotprice
			end if
			
			set rs=server.createobject("ADODB.RecordSet")
			sql="select * from competitive"
			rs.open sql,conn,1,3
			rs.addnew
			rs("djh")=nowdjh
			rs("tenderid")=tenderids(i)
			rs("companyid")=nowcompanyid
			rs("singleprice")=nowsingleprice
			rs("totalprice")=nowtotalprice
			rs("spotprice")=nowspotprice
			rs("addtime")=now()
			rs("statue")=0
			rs("ifzu")=0
			rs.update
			rs.close
			
			alltotalprice=alltotalprice+0+nowtotalprice
			allspotprice=allspotprice+0+nowspotprice
			
			response.write "<font color='#FF0000'>"&shownameint("material","tender","id",tenderids(i))&"</font> 报价成功，总价 <font color='#0000FF'>"&nowtotalprice&"</font>，到现场价 <font color='#0000FF'>"&nowspotprice&"</font>！<br>"
			response.flush
		next
		
		set rsm=server.createobject("ADODB.RecordSet")
		sqlm="select * from competitive"
		rsm.open sqlm,conn,1,3
		rsm.addnew
		rsm("djh")=nowdjh
		rsm("tenderid")=0
		rsm("companyid")=nowcompanyid
		rsm("singleprice")=0
		rsm("totalprice")=alltotalprice
		rsm("spotprice")=allspotprice
		rsm("deliverydate")=nowdeliverydate
		rsm("paystyle")=nowpaystyle
		rsm("beizhu")=nowbeizhu
		rsm("addtime")=now()
		rsm("statue")=0
		rsm("ifzu")=1
		rsm.update
		rsm.close
			
		response.write "<br>总记录生成成功，总价 <font color='#0000FF'>"&alltotalprice&"</font>，总到现场价 <font color='#0000FF'>"&allspotprice&"</font>！<br>"
		response.flush
		
		set rsc=conn.execute("select id from competitive where djh='"&nowdjh&"' and companyid="&nowcompanyid&" and ifzu=1")
		if rsc.eof then
			If IsSqlDataBase=0 Then
				conn.execute("insert into competitive (djh,tenderid,companyid,singleprice,totalprice,spotprice,deliverydate,paystyle,beizhu,addtime,statue,ifzu) values ('"&nowdjh&"',0,"&nowcompanyid&",0,"&alltotalprice&","&allspotprice&",#"&nowdeliverydate&"#,'"&nowpaystyle&"','"&nowbeizhu&"',#"&now()&"#,0,1)")
			else
				conn.execute("insert into competitive (djh,tenderid,companyid,singleprice,totalprice,spotprice,deliverydate,paystyle,beizhu,addtime,statue,ifzu) values ('"&nowdjh&"',0,"&nowcompanyid&",0,"&alltotalprice&","&allspotprice&",'"&nowdeliverydate&"','"&nowpaystyle&"','"&nowbeizhu&"','"&now()&"',0,1)")
			end if
			
			response.write "<br>总记录生成成功，总价 <font color='#0000FF'>"&alltotalprice&"</font>，总到现场价 <font color='#0000FF'>"&allspotprice&"</font>！<br>"
			response.flush
		end if
		rsc.close
		
		call HintAndTurn("竞价信息添加成功，请注意关注您是否中标！",request.form("fromurl")&".asp?"&gourl)
		response.end
	end if
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'beizhu',
cssPath : 'edit/index.css'
});
function checkenter(){
	if (document.form.deliverydate.value == ""){
		alert("请填写交货日期！");
		form.deliverydate.focus();
		return false;
	}
	if (document.form.paystyle.value == ""){
		alert("请填写付款方式！");
		form.paystyle.focus();
		return false;
	}
}
function jisuan(sid)
{
var sid=sid;
var shuliang;
var singleprice;
shuliang=0;
singleprice=0;
shuliang=eval("document.form.material_shuliang"+ sid +".value");
singleprice=eval("document.form.singleprice"+ sid +".value");
if(shuliang!="" && shuliang>0 && singleprice!="" && singleprice>=0)
{
	eval("document.form.totalprice"+ sid +".value=shuliang*singleprice");
	eval("document.form.spotprice"+ sid +".value=shuliang*singleprice");
}
else
{
	eval("document.form.totalprice"+ sid +".value=''");
	eval("document.form.spotprice"+ sid +".value=''");
}
}
</script>
<%
set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
	if (rs_t("startdate")=date() and rs_t("startdatehour")>hour(now()) or rs_t("startdate")>date()) then
		call HintAndBack("很抱歉，尚未到竞价开始时间！",1)
	elseif (rs_t("enddate")=date() and rs_t("enddatehour")<=hour(now()) or rs_t("enddate")<date()) then
		call HintAndBack("很抱歉，竞价截止时间已过！",1)
	else
%>
<form name='form' method="post">
<input type="hidden" name="action" value="saveadd">
<input type="hidden" name="companyid" value="<%=session("user_id")%>">
<input type="hidden" name="id" value="<%=request("id")%>">
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
<input type="hidden" name="checkstatue" value="<%=request("checkstatue")%>">
<input type="hidden" name="checkcondition" value="<%=request("checkcondition")%>">
<input type="hidden" name="checkstyle" value="<%=request("checkstyle")%>">
<input type="hidden" name="keyword1" value="<%=request("keyword1")%>">
<input type="hidden" name="page" value="<%=request("page")%>">
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0' style="word-break: break-all;">
<tr>
	<td class="grid_header" colspan="4"><strong>基本信息</strong></td>
</tr>
<tr>
	<td class="list_header_required" width="15%">项目名称</td>
	<td class="list_required" width="35%"><%=rs_t("tendername")%></td>
	<td class="list_header_required" width="15%">竞价编号</td>
	<td class="list_required" width="35%"><%=rs_t("djh")%></td>
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
	<td class="grid_header" colspan="4"><strong>交货和付款</strong></td>
</tr>
<tr>
	<td class="list_header_required">交货时间</td>
	<td class="list_required" colspan="3"><input type='text' class='form_textbox_normal' name='deliverydate' id='deliverydate' value='' size='20' onFocus="calendar()"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">付款方式</td>
	<td class="list_required" colspan="3"><input type='text' class='form_textbox_normal' name='paystyle' id='paystyle' value='' size='80'><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="grid_header" colspan="4"><strong>采购物资报价</strong></td>
</tr>
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs_t("djh")&"'")
if not rs_s.eof then
	wzno=0
	do while not rs_s.eof
	wzno=wzno+1
	response.write "<tr>"%>
	<td class="list_header_required">物资<%=wzno%></td>
	<td class="list_required" >
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required" height="34">名称：<%=rs_s("material")%></td>
		</tr>
		<tr>
			<td class="list_required" height="34">规格：<%=rs_s("material_guige")%></td>
		</tr>
		<tr>
			<td class="list_required" height="34">材质：<%=rs_s("material_caizhi")%></td>
		</tr>
		<tr>
			<td class="list_required" height="34">单位：<%=rs_s("material_danwei")%></td>
		</tr>
		</table>
	</td>
	<td class="list_required" colspan="2">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<input type="hidden" name="tenderid" value="<%=rs_s("id")%>">
		<input type="hidden" name="material_shuliang<%=rs_s("id")%>" value="<%=rs_s("material_shuliang")%>">
		<tr>
			<td class="list_required" height="34"><strong>物资数量：<%=rs_s("material_shuliang")%></strong></td>
		</tr>
		<tr>
			<td class="list_required">物资单价：<input type='text' class='form_textbox_normal' name='singleprice<%=rs_s("id")%>' id='singleprice<%=rs_s("id")%>' value='' size='20' onKeyUp="value=value.replace(/[^\d.]/g,'');jisuan('<%=rs_s("id")%>');" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d.]/g,''))">元</td>
		</tr>
		<tr>
			<td class="list_required">物资总价：<input type='text' readonly class='form_textbox_normal' name='totalprice<%=rs_s("id")%>' id='totalprice<%=rs_s("id")%>' value='' size='20' onKeyUp="value=value.replace(/[^\d.]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d.]/g,''))">元</td>
		</tr>
		<tr>
			<td class="list_required">到现场价：<input type='text' class='form_textbox_normal' name='spotprice<%=rs_s("id")%>' id='spotprice<%=rs_s("id")%>' value='' size='20' onKeyUp="value=value.replace(/[^\d.]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d.]/g,''))">元</td>
		</tr>
		</table>
	</td>
	<%response.write "</tr>"
	rs_s.movenext
	loop
end if
rs_s.close%>
<tr>
	<td class="grid_header" colspan="4"><strong>备注</strong></td>
</tr>
<tr>
	<td class="list_header_required">备注</td>
	<td class="list_required" colspan="3"><textarea id="beizhu" name="beizhu" style="width:620px;height:150px;visibility:hidden;"></textarea></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="3"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 确定 ' onClick="return checkenter();">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</table>
</form>
<%	end if
end if
rs_t.close%>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->