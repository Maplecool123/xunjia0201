<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
call WhereTable("identity.png","修改竞价信息")

gourl="needtime1="&request.form("needtime1")&"&needtime2="&request.form("needtime2")&"&needtime3="&request.form("needtime3")&"&needtime4="&request.form("needtime4")&"&start_date="&request.form("start_date")&"&end_date="&request.form("end_date")&"&tender_class="&request.form("tender_class")&"&checkgrade="&request.form("checkgrade")&"&checkstyle="&request.form("checkstyle")&"&checkcondition="&request.form("checkcondition")&"&checkstatue="&request.form("checkstatue")&"&keyword1="&request.form("keyword1")&"&page="&request.form("page")

if request.form("action")="saveadd" then
	addstr=""
	nowtendername=RemoveHTML(ReplaceQM(request.form("tendername")))
	nowdjh=request.form("djh")
	nowbasis=request.form("basis")
	nowtenderclass=request.form("tenderclass")
	nowneeddate=request.form("needdate")
	nowstartdate=request.form("startdate")
	nowstartdatehour=request.form("startdatehour")
	nowenddate=request.form("enddate")
	nowenddatehour=request.form("enddatehour")
	nowdetail=request.form("detail")
	nowcompanyno=checkifnum(request.form("companyno"))
	nowminmoney=checkifnum(request.form("minmoney"))
	nowtendergradeid=checkifnum(request.form("tendergradeid"))
	if nowtendergradeid<>0 then
		nowtendergrade=shownameint("companygradeorder","companygrade","id",nowtendergradeid)
	else
		nowtendergrade=0
	end if
	
	nowmaterialno=checkifnum(request.form("materialno"))
	if nowmaterialno-maxsubtender<=0 then nowmaterialno=maxsubtender
	
	call warnifdate(nowneeddate,"需求日期")
	call warnifdate(nowstartdate,"报价开始时间")
	call warnifdate(nowenddate,"报价结束时间")
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from tender where djh='"&nowdjh&"'"
	rs.open sql,conn,1,3
	if not rs.eof then
		nowdjh=getlastdjhno(date(),IsSqlDataBase)
		addstr=addstr&"\n\n由于编号重复，本竞价信息编号更改为"&nowdjh
	end if
	rs.close
	
	'conn.execute("delete from [tender] where djh='"&request.form("lastdjh")&"'")
	'conn.execute("delete from [competitive] where djh='"&request.form("lastdjh")&"'")
	conn.execute("update [tender] set statue=4 where djh='"&request.form("lastdjh")&"'")
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from tender"
	rs.open sql,conn,1,3
	rs.addnew
	rs("tendername")=nowtendername
	rs("djh")=nowdjh
	rs("lastdjh")=request.form("lastdjh")
	rs("basis")=nowbasis
	rs("tenderclass")=nowtenderclass
	rs("needdate")=nowneeddate
	rs("startdate")=nowstartdate
	rs("startdatehour")=nowstartdatehour
	rs("enddate")=nowenddate
	rs("enddatehour")=nowenddatehour
	rs("detail")=nowdetail
	rs("tendergrade")=nowtendergrade
	rs("tendergradeid")=nowtendergradeid
	rs("companyno")=nowcompanyno
	rs("minmoney")=nowminmoney
	rs("addman")=session("user_id")
	rs("addtime")=now()
	rs("finalcompany")=0
	rs("statue")=0
	rs("ifzu")=1
	rs("ifdel")=0
	rs.update
	rs.close
	
	for x=1 to nowmaterialno
	nowmaterial=request.form("material"&x)
	nowmaterial_guige=request.form("material_guige"&x)
	nowmaterial_caizhi=request.form("material_caizhi"&x)
	nowmaterial_danwei=request.form("material_danwei"&x)
	nowmaterial_shuliang=checkifnum(request.form("material_shuliang"&x))
	if nowmaterial<>"" then
		if nowmaterial_shuliang<=0 then
			addstr=addstr&"\n\n第"&x&"项物资，数量不可以为0或者小于零，请核实。"
		else
			set rs=server.createobject("ADODB.RecordSet")
			sql="select * from [tender]"
			rs.open sql,conn,1,3
			rs.addnew
			rs("djh")=nowdjh
			rs("material")=nowmaterial
			rs("material_guige")=nowmaterial_guige
			rs("material_caizhi")=nowmaterial_caizhi
			rs("material_danwei")=nowmaterial_danwei
			rs("material_shuliang")=nowmaterial_shuliang
			rs("addman")=session("user_id")
			rs("addtime")=now()
			rs("finalcompany")=0
			rs("statue")=0
			rs("ifzu")=0
			rs("ifdel")=0
			rs.update
			rs.close
		end if
	end if
	next
	
	call HintAndTurn("竞价信息发布成功！"&addstr,"admin_tender_float.asp?"&gourl)
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'detail',
cssPath : 'edit/index.css'
});

function checkenter(){
	var ifmaterial;
	ifmaterial="no";
	if (document.form.tendername.value == ""){
		alert("请填写项目名称！");
		form.tendername.focus();
		return false;
	}
	if (document.form.djh.value == ""){
		alert("请填写竞价编号！");
		form.djh.focus();
		return false;
	}
	if (document.form.basis.value == ""){
		alert("请填写采购依据！");
		form.basis.focus();
		return false;
	}
	if (document.form.needdate.value == ""){
		alert("请填写需求日期！");
		form.needdate.focus();
		return false;
	}
	if (document.form.startdate.value == ""){
		alert("请填写报价开始时间！");
		form.startdate.focus();
		return false;
	}
	if (document.form.enddate.value == ""){
		alert("请填写截止开始时间！");
		form.enddate.focus();
		return false;
	}
	if (document.form.detail.value == ""){
		alert("请填写详情说明！");
		return false;
	}
	if (document.form.companyno.value == ""){
		alert("请填写竞价成功条件之竞价企业数量！");
		form.companyno.focus();
		return false;
	}
	if (document.form.minmoney.value == ""){
		alert("请填写竞价成功条件之竞价最高金额！");
		form.minmoney.focus();
		return false;
	}
	
	for (i=1;i<=<%=maxsubtender%>;i++)
	{
	if (eval("document.form.material"+ i +".value != ''")){
		ifmaterial="yes";
		if (eval("document.form.material_shuliang"+ i +".value == '' || document.form.material_shuliang"+ i +".value<=0"))
		{
			eval("alert('请输入第"+ i +"项物资采购数量，并且必须大于零！')");
			eval("form.material_shuliang"+ i +".focus()");
			return false;
		}
	}
	}
	
	if (ifmaterial == "no"){
		alert("请填写至少一项竞价物资！");
		form.material1.focus();
		return false;
	}
}

var xmlHttp = false;
try {
  xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
} catch (e) {
  try {
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  } catch (e2) {
    xmlHttp = false;
  }
}
if (!xmlHttp && typeof XMLHttpRequest != 'undefined') {
  xmlHttp = new XMLHttpRequest();
}

function callServer1() {
  var p_basis = document.getElementById("basis").value;
  if ((p_basis == null) || (p_basis == "")) return;
  var url = "Checkbasis.asp?basis=" + escape(p_basis);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage1;
  xmlHttp.send(null);  
}

function updatePage1() {
  if (xmlHttp.readyState < 4) {
    test1.innerHTML="正在查询该采购依据是否已存在";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
    test1.innerHTML=response;
  }
}

function callServer2() {
  var p_tendername = document.getElementById("tendername").value;
  if ((p_tendername == null) || (p_tendername == "")) return;
  var url = "Checktender.asp?tendername=" + escape(p_tendername);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage2;
  xmlHttp.send(null);  
}

function updatePage2() {
  if (xmlHttp.readyState < 4) {
    test2.innerHTML="正在查询该用户名是否已存在";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
    test2.innerHTML=response;
  }
}

function firm(firmtxt,firmno){
if(confirm(firmtxt)){
	eval("document.form.material"+ firmno +".value=''");
	eval("document.form.material_guige"+ firmno +".value=''");
	eval("document.form.material_caizhi"+ firmno +".value=''");
	eval("document.form.material_danwei"+ firmno +".value=''");
	eval("document.form.material_shuliang"+ firmno +".value=''");
	eval("document.all.shenqing"+ firmno +".style.display='none';");
}
}
</script>
<%
set rs_t=conn.execute("select * from tender where id="&request("id")&"")
if not rs_t.eof then
%>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form' method="post">
<input type="hidden" name="action" value="saveadd">
<input type="hidden" name="id" value="<%=rs_t("id")%>">
<input type="hidden" name="lastdjh" value="<%=rs_t("djh")%>">
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
<input type="hidden" name="checkgrade" value="<%=request("checkgrade")%>">
<input type="hidden" name="checkcondition" value="<%=request("checkcondition")%>">
<input type="hidden" name="checkstyle" value="<%=request("checkstyle")%>">
<input type="hidden" name="keyword1" value="<%=request("keyword1")%>">
<input type="hidden" name="page" value="<%=request("page")%>">
<tr>
	<td class="list_header_required">项目名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='tendername' value='<%=rs_t("tendername")%>' onKeyUp="callServer2();" style="width:300px;"><font color="#FF0000">*</font></td>
	<td class="list_required"><span id="test2"></span></td>
</tr>
<tr>
	<td class="list_header_required">竞价编号</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='djh' id='djh' value='<%=getlastdjhno(date(),IsSqlDataBase)%>' style="width:300px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required" width="15%">采购依据</td>
	<td class="list_required" width="45%"><input type='text' class='form_textbox_normal' name='basis' id='basis' value='<%=rs_t("basis")%>' onKeyUp="callServer1();" style="width:300px;"><font color="#FF0000">*</font></td>
	<td class="list_required" width="40%"><span id="test1"></span></td>
</tr>
<tr>
	<td class="list_header_required">项目类型</td>
	<td class="list_required" colspan="2">材料<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="材料"<%if rs_t("tenderclass")="材料" then response.write " checked"%>> 设备<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="设备"<%if rs_t("tenderclass")="设备" then response.write " checked"%>> 综合<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="综合"<%if rs_t("tenderclass")="综合" then response.write " checked"%>><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">竞价等级</td>
	<td class="list_required" colspan="2"><select class='form_combo_normal' name='tendergradeid'>
	<option value="0"<%if rs_t("tendergrade")=0 then response.write " selected"%>>所有等级</option>
	<%set rsa=conn.execute("select * from companygrade order by companygradeorder")
	if not rsa.eof then
		do while not rsa.eof
		response.write "<option value="&rsa("id")
		if rsa("companygradeorder")-checkifnum(rs_t("tendergrade"))=0 then response.write " selected"
		response.write ">"&rsa("companygradename")&"</option>"
		rsa.movenext
		loop
	end if
	rsa.close%>
	</select><font color="#FF0000">*</font>
	（点击“<a title=竞价预期总额说明 onMouseOver="window.status='竞价等级说明';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("竞价等级设定后，等于或者高于这个级别的企业方可参加竞价。")>这里</a>”查看详细说明。）</td>
</tr>
<tr>
	<td class="list_header_required">需求日期</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='needdate' value='<%=rs_t("needdate")%>' onfocus="calendar()" style="width:100px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">报价开始时间</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='startdate' value='<%=rs_t("startdate")%>' onfocus="calendar()" style="width:100px;">
	<select class='form_combo_normal' name="startdatehour">
	<%for i=0 to 23
		response.write "<option value="&i
		if i-rs_t("startdatehour")=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>点<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">报价截止时间</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='enddate' value='<%=rs_t("enddate")%>' onfocus="calendar()" style="width:100px;">
	<select class='form_combo_normal' name="enddatehour">
	<%for i=0 to 23
		response.write "<option value="&i
		if i-rs_t("enddatehour")=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>点<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">详情说明</td>
	<td class="list_required" colspan="2"><textarea id="detail" name="detail" style="width:620px;height:150px;visibility:hidden;"><%=rs_t("detail")%></textarea></td>
</tr>
<%set rs_s=conn.execute("select * from tender where ifzu=0 and ifdel=0 and djh='"&rs_t("djh")&"'")
if not rs_s.eof then
	x=0
	do while not rs_s.eof
	x=x+1%>
<tr id="shenqing<%=x%>">
	<td class="list_header_required">物资<%=x%></td>
	<td class="list_required">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required">名称：<input type='text' class='form_textbox_normal' name='material<%=x%>' id='material<%=x%>' value='<%=rs_s("material")%>' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">规格：<input type='text' class='form_textbox_normal' name='material_guige<%=x%>' id='material_guige<%=x%>' value='<%=rs_s("material_guige")%>' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">材质：<input type='text' class='form_textbox_normal' name='material_caizhi<%=x%>' id='material_caizhi<%=x%>' value='<%=rs_s("material_caizhi")%>' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">单位：<input type='text' class='form_textbox_normal' name='material_danwei<%=x%>' id='material_danwei<%=x%>' value='<%=rs_s("material_danwei")%>' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">数量：<input type='text' class='form_textbox_normal' name='material_shuliang<%=x%>' id='material_shuliang<%=x%>' value='<%=rs_s("material_shuliang")%>' style="width:300px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
		</tr>
		</table>
	</td>
	<td class="list_required">
	<%if x<maxsubtender then%><input type="button" value="下一个" onClick="shenqing<%=(x+1)%>.style.display='';" class='form_button'>&nbsp;&nbsp;<%end if%>
	<%if x<>1 then%><input type="button" value="清除" onClick="firm('确定清除第 <%=x%> 条么？',<%=x%>);" class='form_submit'><%end if%>
	</td>
</tr>		
	<%rs_s.movenext
	loop
end if
rs_s.close%>
<input type="hidden" name="materialno" value="<%=x%>">
<%for y=x+1 to maxsubtender%>
<tr id="shenqing<%=y%>"<%if y<>1 then%> style="display:none;"<%end if%>>
	<td class="list_header_required">物资<%=y%></td>
	<td class="list_required">
		<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
		<tr>
			<td class="list_required">名称：<input type='text' class='form_textbox_normal' name='material<%=y%>' id='material<%=y%>' value='' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">规格：<input type='text' class='form_textbox_normal' name='material_guige<%=y%>' id='material_guige<%=y%>' value='' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">材质：<input type='text' class='form_textbox_normal' name='material_caizhi<%=y%>' id='material_caizhi<%=y%>' value='' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">单位：<input type='text' class='form_textbox_normal' name='material_danwei<%=y%>' id='material_danwei<%=y%>' value='' style="width:300px;"></td>
		</tr>
		<tr>
			<td class="list_required">数量：<input type='text' class='form_textbox_normal' name='material_shuliang<%=y%>' id='material_shuliang<%=y%>' value='' style="width:300px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
		</tr>
		</table>
	</td>
	<td class="list_required">
	<%if y<>maxsubtender then%><input type="button" value="下一个" onClick="shenqing<%=(y+1)%>.style.display='';" class='form_button'>&nbsp;&nbsp;<%end if%>
	<%if y<>1 then%><input type="button" value="清除" onClick="firm('确定清除第 <%=y%> 条么？',<%=y%>);" class='form_submit'><%end if%>
	</td>
</tr>		
<%next%>
<tr>
	<td class="list_header_required">竞价成功条件</td>
	<td class="list_required" colspan="2">竞价企业数量达到<input type='text' class='form_textbox_normal' name='companyno' value='<%=rs_t("companyno")%>' style="width:100px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">家<font color="#FF0000">*</font><br>
	竞价最高金额达到<input type='text' class='form_textbox_normal' name='minmoney' value='<%=rs_t("minmoney")%>' style="width:100px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">元<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 确定并再次发布 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<%end if
rs_t.close%>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->