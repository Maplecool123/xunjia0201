<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
	gourl="needtime1="&request("needtime1")&"&needtime2="&request("needtime2")&"&needtime3="&request("needtime3")&"&needtime4="&request("needtime4")&"&start_date="&request("start_date")&"&end_date="&request("end_date")&"&tender_class="&request("tender_class")&"&checkgrade="&request("checkgrade")&"&checkstyle="&request("checkstyle")&"&checkcondition="&request("checkcondition")&"&checkstatue="&request("checkstatue")&"&keyword1="&request("keyword1")&"&page="&request("page")
	
	useridlikestr=","&session("user_id")&","
	nullstr=""
	
	set rs=server.createobject("ADODB.RecordSet")
	sql="select focusman from tender where id="&request("id")&""
	rs.open sql,conn,1,3
	if not rs.eof then
		if rs(0)=useridlikestr then
			rs("focusman")=nullstr
			rs.update
		else
			if instr(rs(0),useridlikestr)>0 then
				rs("focusman")=replace(rs("focusman"),useridlikestr,",")
				rs.update
			end if
		end if
	end if
	rs.close
	
	call HintAndTurn("取消收藏成功！","tender_mycollection.asp?"&gourl&"")
else
	response.redirect "erro.asp"
	response.end
end if%>
<!--#include file="Index_bottom.asp" -->