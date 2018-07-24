<!--#include file="Inc/Function.asp" -->
<%
dim weekNoNow
weekNoNow = weekNo()
%>
<html>
<head>
<LINK REL=STYLESHEET TYPE="text/css" HREF="css/menu.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body leftmargin=0 rightmargin=0 marginwidth=0 marginheigh=0 topmargin=0 background="images/weekbg.png">
  <div class='week'><%=weekNoNow%></div>  <div class="range">Month - <font color=yellow><b><%=month(now())%></b></font>
  <br />
  <font color=#AAAAAA><%=showDate(true,weekNoNow)%></font>
  <br />
  <font color=#AAAAAA><%=showDate(false,weekNoNow)%></font>
  
  </div>
</body>
</html>
