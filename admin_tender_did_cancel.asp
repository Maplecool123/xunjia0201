<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-1=0 or session("iflogin")-99=0 then
	gourl="needtime1="&request("needtime1")&"&needtime2="&request("needtime2")&"&needtime3="&request("needtime3")&"&needtime4="&request("needtime4")&"&start_date="&request("start_date")&"&end_date="&request("end_date")&"&tender_class="&request("tender_class")&"&checkgrade="&request("checkgrade")&"&checkstyle="&request("checkstyle")&"&checkcondition="&request("checkcondition")&"&checkstatue="&request("checkstatue")&"&keyword1="&request("keyword1")&"&page="&request("page")
	
	djh=shownameint("djh","tender","id",request("id"))
	
	conn.execute("update tender set ifdel=1 where djh='"&djh&"'")
	conn.execute("delete from competitive where djh='"&djh&"'")
	
	call HintAndTurn("取消项竞价成功！",request("fromurl")&".asp?"&gourl&"")
else
	response.redirect "erro.asp"
	response.end
end if%>
<!--#include file="Index_bottom.asp" -->