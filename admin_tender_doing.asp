<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-3=0 or session("iflogin")-99=0 then
'全局变量
Dim CurrentPage,sql,i,rs
'全局变量

call WhereTable("filefind.png","正在竞价")
call tendergotorightstatue

dim nowcheckstyle,nowkeyword1
currentpage=request("page")
'if currentpage="" then currentpage=0
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

if request("action")="delall" then
	id=replace(request("id")," ","")

	goon=0
	if replace(id,",","")<>"" then goon=1

	if goon=0 then
		call HintAndBack("请至少选中一条记录！",1)
	end if
	
	id=split(id,",")
	for i=0 to UBound(id)
		conn.execute("delete from [tender] where djh='"&shownameint("djh","tender","id",id(i))&"'")
	next
	
	call HintAndTurn("删除成功！","?"&gourl&"&page="&currentpage&"")
end if

fromurl="admin_tender_doing"%>
<!--#include file="admin_tender_search_top.asp" -->
<%
dim thisClass
thisClass = "grid_even"

Set Rs = server.createobject("adodb.recordset")
sql="select * from tender where ifdel=0 and ifzu=1 and statue=1"
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
	sql=sql&" and statue="&nowcheckstatue&""
end if
if nowcheckgrade<>"" then
	sql=sql&" and tendergradeid="&nowcheckgrade&""
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
<form name='form1'>
<input type="hidden" name="action" value="delall">
<input type="hidden" name="needtime1" value="<%=nowneedtime1%>">
<input type="hidden" name="needtime2" value="<%=nowneedtime2%>">
<input type="hidden" name="needtime3" value="<%=nowneedtime3%>">
<input type="hidden" name="needtime4" value="<%=nowneedtime4%>">
<input type="hidden" name="start_date" value="<%=nowstart_date%>">
<input type="hidden" name="end_date" value="<%=nowend_date%>">
<input type="hidden" name="company_class" value="<%=nowcompany_class%>">
<input type="hidden" name="company_type" value="<%=nowcompany_type%>">
<input type="hidden" name="checkstatue" value="<%=nowcheckstatue%>">
<input type="hidden" name="checkcondition" value="<%=nowcheckcondition%>">
<input type="hidden" name="checkstyle" value="<%=nowcheckstyle%>">
<input type="hidden" name="keyword1" value="<%=nowkeyword1%>">
<table class='grid_table' cellpadding='0' cellspacing='1' border='0'>
<tr class='grid_header'>
	<td><img src="images/folder_close.gif" style="cursor:hand" onClick="collapseall(this)" />ID号</td>
	<td>项目名称</td>
	<td>竞价编号</td>
	<td>采购依据</td>
	<td>项目类型</td>
	<td>竞价等级</td>
	<td>需求日期</td>
	<td>竞价时间</td>
	<td>发布人</td>
	<td>发布时间</td>
	<td>当前状态</td>
	<td>竞价企业</td>
	<td>详情</td>
	<%if session("iflogin")-1=0 or session("iflogin")-99=0 then%>
	<td>删除</td>
	<%end if%>
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
		<td class="grid_cell"><%=rs("basis")%></td>
		<td class="grid_cell"><%=rs("tenderclass")%></td>
		<td class="grid_cell"><%=showtendergrade(rs("tendergradeid"))%></td>
		<td class="grid_cell"><%=rs("needdate")%></td>
		<td class="grid_cell"><%response.write rs("startdate")&" "&rs("startdatehour")&"点 至 "&rs("enddate")&" "&rs("enddatehour")&"点"%></td>
		<td class="grid_cell"><%=shownameint("userrealname","login","id",rs("addman"))%></td>
		<td class="grid_cell"><%=rs("addtime")%></td>
		<td class="grid_cell"><%=gettenderstatue(rs("statue"))%></td>
		<td class="grid_cell"><%response.write "<font color='#0000FF'><strong>"&rs("companyno")&"</strong></font> / <font color='#FF0000'><strong>"&getmaxcompetitivecompanyno(rs("djh"),"")&"</strong></font>"%></td>
		<td class="grid_cell"><input type='button' value='查看'  onClick="doCheckDetail('admin_tender_main.asp?id=<%=rs("id")%>',<%=Modalwidth%>,<%=Modalheight%>)" class='form_button'></td>
		<%if session("iflogin")-1=0 or session("iflogin")-99=0 then%>
		<td class="grid_cell">
		<%if gettotalitem("tender","djh",rs("djh")," and ifzu=0")<=0 and rs("addman")-session("user_id")=0 then%>
		<input type="checkbox" name="ID" value="<%=rs("id")%>" class="form_checkbox_normal" title="这是一条错误的竞价信息，可以删除">
		<%else
			response.write "----"
		end if%>
		</td>
		<%end if%>
	</tr>
	<tr id="tender<%=xx%>" style="display:none;">
		<td colspan="16" width="100%" align="center"><!--#include file="tender_sub_detail.asp" --></td>
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
				<%showdel="yes"%>
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
</form>
<%else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->