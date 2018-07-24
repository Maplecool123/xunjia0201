<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
call WhereTable("gnome-ccdesktop.png","导入竞价信息第一步：填写竞价基本信息")

if request("action")="saveadd" then
	addstr=""
	nowtendername=RemoveHTML(ReplaceQM(request("tendername")))
	nowdjh=request("djh")
	nowbasis=request("basis")
	nowtenderclass=request("tenderclass")
	nowneeddate=request("needdate")
	nowstartdate=request("startdate")
	nowstartdatehour=request("startdatehour")
	nowenddate=request("enddate")
	nowenddatehour=request("enddatehour")
	nowdetail=request("detail")
	nowcompanyno=checkifnum(request("companyno"))
	nowminmoney=checkifnum(request("minmoney"))
	nowtendergradeid=checkifnum(request("tendergradeid"))
	if nowtendergradeid<>0 then
		nowtendergrade=shownameint("companygradeorder","companygrade","id",nowtendergradeid)
	else
		nowtendergrade=0
	end if
	
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
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select * from tender"
	rs.open sql,conn,1,3
	rs.addnew
	rs("tendername")=nowtendername
	rs("djh")=nowdjh
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
	
	call HintAndTurn("竞价基本信息填写成功，请导入采购物资信息！"&addstr,"admin_tender_add_excelin2.asp?djh="&nowdjh)
end if
%>
<script type="text/javascript" charset="utf-8" src="edit/kindeditor.js"></script>
<script type="text/javascript">
KE.show({
id : 'detail',
cssPath : 'edit/index.css'
});

function checkenter(){
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
</script>
<table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
<form name='form'>
<input type="hidden" name="action" value="saveadd">
<tr>
	<td class="list_header_required">填写竞价</td>
	<td class="list_required" colspan="3"><input class='form_submit' type='button' name='excelin' id='excelin' value=' 进入填写页面 ' onClick="location.href='admin_tender_add.asp'"></td>
</tr>
<tr>
	<td class="list_header_required">项目名称</td>
	<td class="list_required"><input type='text' class='form_textbox_normal' name='tendername' value='' onKeyUp="callServer2();" style="width:300px;"><font color="#FF0000">*</font></td>
	<td class="list_required"><span id="test2"></span></td>
</tr>
<tr>
	<td class="list_header_required">竞价编号</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='djh' id='djh' value='<%=getlastdjhno(date(),IsSqlDataBase)%>' style="width:300px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required" width="15%">采购依据</td>
	<td class="list_required" width="45%"><input type='text' class='form_textbox_normal' name='basis' id='basis' value='' onKeyUp="callServer1();" style="width:300px;"><font color="#FF0000">*</font></td>
	<td class="list_required" width="40%"><span id="test1"></span></td>
</tr>
<tr>
	<td class="list_header_required">项目类型</td>
	<td class="list_required" colspan="2">材料<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="材料" checked> 设备<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="设备"> 综合<input type="radio" class='form_radio_normal' name="tenderclass" id="tenderclass" value="综合"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">竞价等级</td>
	<td class="list_required" colspan="2"><select class='form_combo_normal' name='tendergradeid'>
	<option value="0">所有等级</option>
	<%set rsa=conn.execute("select * from companygrade order by companygradeorder")
	if not rsa.eof then
		do while not rsa.eof
		response.write "<option value="&rsa("id")&">"&rsa("companygradename")&"</option>"
		rsa.movenext
		loop
	end if
	rsa.close%>
	</select><font color="#FF0000">*</font>
	（点击“<a title=竞价预期总额说明 onMouseOver="window.status='竞价等级说明';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("竞价等级设定后，等于或者高于这个级别的企业方可参加竞价。")>这里</a>”查看详细说明。）</td>
</tr>
<tr>
	<td class="list_header_required">需求日期</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='needdate' value='' onfocus="calendar()" style="width:100px;"><font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">报价开始时间</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='startdate' value='<%=date()%>' onfocus="calendar()" style="width:100px;">
	<select class='form_combo_normal' name="startdatehour">
	<%for i=0 to 23
		response.write "<option value="&i
		if i-8=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>点<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">报价截止时间</td>
	<td class="list_required" colspan="2"><input type='text' class='form_textbox_normal' name='enddate' value='' onfocus="calendar()" style="width:100px;">
	<select class='form_combo_normal' name="enddatehour">
	<%for i=0 to 23
		response.write "<option value="&i
		if i-8=0 then response.write " selected"
		response.write ">"&i&"</option>"
	next%>
	</select>点<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_required">详情说明</td>
	<td class="list_required" colspan="2"><textarea id="detail" name="detail" style="width:620px;height:150px;visibility:hidden;"></textarea></td>
</tr>
<tr>
	<td class="list_header_required">竞价成功条件</td>
	<td class="list_required" colspan="2">竞价企业数量达到<input type='text' class='form_textbox_normal' name='companyno' value='1' style="width:100px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">家<font color="#FF0000">*</font><br>
	竞价最高金额达到<input type='text' class='form_textbox_normal' name='minmoney' value='0' style="width:100px;" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">元<font color="#FF0000">*</font></td>
</tr>
<tr>
	<td class="list_header_command">指令</td>
	<td class="list_command" colspan="2"><input class='form_submit' type='submit' name='cmd' id='cmd' value=' 下一步 ' onClick="return checkenter()">
	<input class='form_submit' type="button" name="back" value=" 返回 " onClick="history.go(-1);"></td>
</tr>
</form>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->