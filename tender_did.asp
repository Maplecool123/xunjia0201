<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("identity.png","尚未中标项目")

dim nowcheckstyle,nowkeyword1
currentpage=request("page")
'if currentpage="" then currentpage=0
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

nowneedtime1=request("needtime1")
nowneedtime2=request("needtime2")
nowneedtime3=request("needtime3")
nowneedtime4=request("needtime4")
nowstart_date=request("start_date") 
if nowstart_date="" then
  nowstart_date=date()-day(date()-1)
end if
nowend_date=request("end_date") 
if nowend_date="" then
  nowend_date=date()
end if

call warnifdate(nowstart_date,"起始时间")
call warnifdate(nowend_date,"结束时间")

nowtender_class = request("tender_class")
nowcheckgrade = request("checkgrade")
nowcheckstyle = request("checkstyle")
nowcheckcondition = request("checkcondition")
nowcheckstatue = request("checkstatue")
nowkeyword1 = RemoveHTML(ReplaceQM(request("keyword1")))

gourl="needtime1="&nowneedtime1&"&needtime2="&nowneedtime2&"&needtime3="&nowneedtime3&"&needtime4="&nowneedtime4&"&start_date="&nowstart_date&"&end_date="&nowend_date&"&tender_class="&nowtender_class&"&checkstyle="&nowcheckstyle&"&checkgrade="&nowcheckgrade&"&checkcondition="&nowcheckcondition&"&checkstatue="&nowcheckstatue&"&keyword1="&nowkeyword1&""
%>
<table class='grid_search' cellpadding='0' cellspacing='1' border='0'>
<form name="form2">
<tr class='grid_header'>
	<td>时间选择</td>
	<td>查询条件</td>
	<td>关键字</td>
	<td>查询</td>
	<td>清除</td>
</tr>
<tr class="grid_odd" align="center">
	<td class="grid_cell">
	开始日期<input type="checkbox" name="needtime1" value="yes"<%if nowneedtime1="yes" then%> checked<%end if%>>
	截止日期<input type="checkbox" name="needtime2" value="yes"<%if nowneedtime2="yes" then%> checked<%end if%>>
	需求日期<input type="checkbox" name="needtime3" value="yes"<%if nowneedtime3="yes" then%> checked<%end if%>>
	发布日期<input type="checkbox" name="needtime4" value="yes"<%if nowneedtime4="yes" then%> checked<%end if%>>
	<input name="start_date" value="<%=nowstart_date%>" onFocus="calendar()" style="width:100px">
	—
	<input name="end_date" value="<%=nowend_date%>" onFocus="calendar()" style="width:100px">
	</td>
	<td class="grid_cell">
	<select name="checkgrade" class='form_combo_normal' onChange="form2.submit();">
	<%response.write "<option value=''>资质级别</option>"
	response.write "<option value='0'"
	if nowcheckgrade="0" then response.write " selected"
	response.write ">所有等级</option>"
	set rsc=conn.execute("select * from companygrade where id<>0 order by companygradeorder")
	if not rsc.eof then
	do while not rsc.eof
	response.write "<option value="&rsc(0)
	if nowcheckgrade=trim(cstr(rsc(0))) then
		response.write " selected"
	end if
	response.write ">"&rsc(1)&"</option>"
	rsc.movenext
	loop
	end if
	rsc.close%>
	</select>
	<select name="checkstatue" class='form_combo_normal' onChange="form2.submit();">
	<option value="">定标状态</option>
	<option value="1"<%if nowcheckstatue="1" then%> selected="selected"<%end if%>>已 定 标</option>
	<option value="2"<%if nowcheckstatue="2" then%> selected="selected"<%end if%>>等待定标</option>
	</select>
	<select name="checkcondition" class='form_combo_normal' onChange="form2.submit();">
	<option value="">查询条件</option>
	<option value="1"<%if nowcheckcondition="1" then%> selected="selected"<%end if%>>按项目名称</option>
	<option value="2"<%if nowcheckcondition="2" then%> selected="selected"<%end if%>>按竞价编号</option>
	<option value="3"<%if nowcheckcondition="3" then%> selected="selected"<%end if%>>按采购依据</option>
	<option value="4"<%if nowcheckcondition="4" then%> selected="selected"<%end if%>>按详情说明</option>
	<option value="5"<%if nowcheckcondition="5" then%> selected="selected"<%end if%>>按物资名称</option>
	<option value="6"<%if nowcheckcondition="6" then%> selected="selected"<%end if%>>按发布人</option>
	</select>
	<select name="checkstyle" class='form_combo_normal' onChange="form2.submit();">
	<option value="mohu"<%if nowcheckstyle="mohu" then%> selected="selected"<%end if%>>模糊</option>
	<option value="jing"<%if nowcheckstyle="jing" then%> selected="selected"<%end if%>>精确</option>
	</select></td>
	<td class="grid_cell"><input type='text' class='form_textbox_normal' name='keyword1' value='<%=nowkeyword1%>' size='15' ></td>
	<td class="grid_cell"><input type='submit' value='查询' class='form_button'></td>
	<td class="grid_cell"><input type='button' value='清除' class='form_button' onClick="window.location.href='?'"></td>
</tr>
</form>
</table>
<%
dim thisClass
thisClass = "grid_even"

dim companyinfo_class,companyinfo_grade_order
call getcompanyinfo(session("user_id"))

Set Rs = server.createobject("adodb.recordset")
sql="select * from tender where ifdel=0 and ifzu=1 and (statue=2 or statue=3 or statue=4)"
'sql=sql&" and pretotalmoney>="&companyinfo_grade_price1&" and pretotalmoney<="&companyinfo_grade_price2&""
sql=sql&" and djh in (select djh from competitive where companyid="&session("user_id")&" and (statue=0 or statue=3) and ifzu=1)"
if companyinfo_class<>"综合" then
	sql=sql&" and (tenderclass='"&companyinfo_class&"' or tenderclass='综合')"
end if
if IsSqlDataBase=0 then
	if nowneedtime1="yes" then
		sql=sql&" and startdate between #"&nowstart_date&"# and #"&nowend_date&"#"
	end if
	if nowneedtime2="yes" then
		sql=sql&" and enddate between #"&nowstart_date&"# and #"&nowend_date&"#"
	end if
	if nowneedtime3="yes" then
		sql=sql&" and needdate between #"&nowstart_date&"# and #"&nowend_date&"#"
	end if
	if nowneedtime4="yes" then
		sql=sql&" and addtime between #"&nowstart_date&"# and #"&cdate(nowend_date)+1&"#"
	end if
else
	if nowneedtime1="yes" then
		sql=sql&" and startdate between '"&nowstart_date&"' and '"&nowend_date&"'"
	end if
	if nowneedtime2="yes" then
		sql=sql&" and enddate between '"&nowstart_date&"' and '"&nowend_date&"'"
	end if
	if nowneedtime3="yes" then
		sql=sql&" and needdate between '"&nowstart_date&"' and '"&nowend_date&"'"
	end if
	if nowneedtime4="yes" then
		sql=sql&" and addtime between '"&nowstart_date&"' and '"&cdate(nowend_date)+1&"'"
	end if
end if
if nowtender_class<>"" then
	sql=sql&" and tenderclass='"&nowtender_class&"'"
end if
if nowcheckstatue<>"" then
	select case nowcheckstatue
	case "1"
		sql=sql&" and finalcompany<>0"
	case "2"
		sql=sql&" and finalcompany=0"
	end select
end if
if nowkeyword1<>"" then
	if nowcheckstyle="mohu" then
		select case nowcheckcondition
		case "1"
		sql=sql&" and tendername like '%"&nowkeyword1&"%'"
		case "2"
		sql=sql&" and djh like '%"&nowkeyword1&"%'"
		case "3"
		sql=sql&" and basis like '%"&nowkeyword1&"%'"
		case "4"
		sql=sql&" and detail like '%"&nowkeyword1&"%'"
		case "5"
		sql=sql&" and djh in (select djh from tender where ifzu=0 and ifdel=0 and material like '%"&nowkeyword1&"%')"
		case "6"
		sql=sql&" and addman in (select id from login where userrealname like '%"&nowkeyword1&"%' or username like '%"&nowkeyword1&"%')"
		case else
		sql=sql&" and (tendername like '%"&nowkeyword1&"%' or djh like '%"&nowkeyword1&"%' or basis like '%"&nowkeyword1&"%' or detail like '%"&nowkeyword1&"%' or djh in (select djh from tender where ifzu=0 and ifdel=0 and material like '%"&nowkeyword1&"%') or addman in (select id from login where userrealname like '%"&nowkeyword1&"%' or username like '%"&nowkeyword1&"%')"
		end select
	elseif nowcheckstyle="jing" then
		select case nowcheckcondition
		case "1"
		sql=sql&" and tendername = '"&nowkeyword1&"'"
		case "2"
		sql=sql&" and djh = '"&nowkeyword1&"'"
		case "3"
		sql=sql&" and basis = '"&nowkeyword1&"'"
		case "4"
		sql=sql&" and detail = '"&nowkeyword1&"'"
		case "5"
		sql=sql&" and djh in (select djh from tender where ifzu=0 and ifdel=0 and material = '"&nowkeyword1&"')"
		case "6"
		sql=sql&" and addman in (select id from login where userrealname = '"&nowkeyword1&"' or username = '"&nowkeyword1&"')"
		case else
		sql=sql&" and (tendername = '"&nowkeyword1&"' or djh = '"&nowkeyword1&"' or basis = '"&nowkeyword1&"' or detail = '"&nowkeyword1&"' or djh in (select djh from tender where ifzu=0 and ifdel=0 and material = '"&nowkeyword1&"') or addman in (select id from login where userrealname = '"&nowkeyword1&"' or username = '"&nowkeyword1&"')"
		end select
	end if
end if
sql=sql&" order by addtime desc"%>
<table class='grid_table' cellpadding='0' cellspacing='1' border='0'>
  <tr class='grid_header'>
    <td><img src="images/folder_close.gif" style="cursor:hand" onClick="collapseall(this)" />ID号</td>
    <td>项目名称</td>
    <td>竞价编号</td>
    <td>项目类型</td>
    <td>需求日期</td>
    <td>竞价时间</td>
    <td>发布时间</td>
    <td>当前状态</td>
    <td>详情</td>
  </tr>
  <%
set rs =server.createobject("ADODB.RecordSet")	
rs.open sql,conn,1,3
if not rs.eof then
	xx=0
	rs.pagesize=maxrecord
	rs.absolutepage=currentpage
	for currentrec=1 to rs.pagesize
    if rs.eof then
		exit for
    end if
	xx=xx+1
	if thisClass = "grid_even" then thisClass = "grid_odd" else thisClass = "grid_even"
	%>
  <tr class=<%=thisClass%> align="center">
    <td class="grid_cell"><img src="images/folder_close.gif" style="cursor:hand" onClick="collapse(this, 'tender<%=xx%>')" /> <%=rs("id")%></td>
    <td class="grid_cell"><font color=green><%=rs("tendername")%></font> </td>
    <td class="grid_cell"><font color=green><%=rs("djh")%></font></td>
    <td class="grid_cell"><%=rs("tenderclass")%></td>
    <td class="grid_cell"><%=rs("needdate")%></td>
    <td class="grid_cell"><%response.write rs("startdate")&" "&rs("startdatehour")&"点 至 "&rs("enddate")&" "&rs("enddatehour")&"点"%></td>
    <td class="grid_cell"><%=rs("addtime")%></td>
    <td class="grid_cell"><%response.write gettenderstatue(rs("statue"))
	if checkifnum(rs("finalcompany"))=0 then response.write "(尚无企业中标)"%></td>
    <td class="grid_cell"><input name="button" type='button' class='form_button'  onClick="doCheckDetail('tender_competitive_main.asp?id=<%=rs("id")%>',<%=Modalwidth%>,<%=Modalheight%>)" value='查看'></td>
  </tr>
  <tr id="tender<%=xx%>" style="display:none;">
    <td colspan="16" width="100%" align="center"><!--#include file="tender_sub_detail_get.asp" --></td>
  </tr>
  <%
	Rs.movenext
	next
	
	if rs.recordcount>0 then 
	%>
  <tr class='grid_header'>
    <td colspan="16">
      <table class='list_table' width='100%' cellpadding='0' cellspacing='1' border='0'>
        <tr class="list_command">
          <td align="right">
            <%showdel="no"%>
            <!--#include file="inc/page_bar.asp" --></td>
        </tr>
    </table></td>
  </tr>
  <%end if
else
	response.write "<tr class='grid_span'><td colspan='16'><font color='red'>查无资料纪录！</font></td></tr>"
end if

rs.close
set rs = nothing
%>
</table>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->