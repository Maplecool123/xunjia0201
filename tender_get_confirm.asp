<!--#include file="Index_top.asp" -->
<%
if session("iflogin")=0 and session("statue")=1 then
	gourl="needtime1="&request("needtime1")&"&needtime2="&request("needtime2")&"&needtime3="&request("needtime3")&"&needtime4="&request("needtime4")&"&start_date="&request("start_date")&"&end_date="&request("end_date")&"&tender_class="&request("tender_class")&"&checkgrade="&request("checkgrade")&"&checkstyle="&request("checkstyle")&"&checkcondition="&request("checkcondition")&"&checkstatue="&request("checkstatue")&"&keyword1="&request("keyword1")&"&page="&request("page")
	
	conn.execute("update competitive set statue=2 where djh='"&request("djh")&"' and companyid="&session("user_id")&"")
	call HintAndTurn("确认中标成功！","tender_get.asp?"&gourl&"")
else
	response.redirect "erro.asp"
	response.end
end if%>
<!--#include file="Index_bottom.asp" -->