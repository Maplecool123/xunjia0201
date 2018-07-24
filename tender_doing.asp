<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("identity.png","正在竞价项目")
call tendergotorightstatue

dim nowcheckstyle,nowkeyword1
currentpage=request("page")
'if currentpage="" then currentpage=0
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if
%>
<!--#include file="tender_search_top.asp" -->
<%
dim thisClass
thisClass = "grid_even"

dim companyinfo_class,companyinfo_grade_order
call getcompanyinfo(session("user_id"))

Set Rs = server.createobject("adodb.recordset")
sql="select * from tender where ifdel=0 and ifzu=1 and statue=1"
'sql=sql&" and pretotalmoney>="&companyinfo_grade_price1&" and pretotalmoney<="&companyinfo_grade_price2&""
sql=sql&" and djh in (select djh from competitive where companyid="&session("user_id")&" and ifzu=1)"
if companyinfo_class<>"综合" then
	sql=sql&" and (tenderclass='"&companyinfo_class&"' or tenderclass='综合')"
end if
if IsSqlDataBase=0 then
	'sql=sql&" and ((startdate=#"&date()&"# and startdatehour<="&hour(now())&") or startdate<#"&date()&"#) and ((enddate=#"&date()&"# and enddatehour>="&hour(now())&") or enddate>#"&date()&"#)"
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
	'sql=sql&" and ((startdate='"&date()&"' and startdatehour<="&hour(now())&") or startdate<'"&date()&"') and ((enddate='"&date()&"' and enddatehour>="&hour(now())&") or enddate>'"&date()&"')"
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
	sql=sql&" and statue="&nowcheckstatue&""
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
		<td class="grid_cell"><%=gettenderstatue(rs("statue"))%></td>
		<td class="grid_cell"><input type='button' value='查看'  onClick="doCheckDetail('tender_competitive_main.asp?id=<%=rs("id")%>',<%=Modalwidth%>,<%=Modalheight%>)" class='form_button'></td>
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
			</table>
		</td>
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