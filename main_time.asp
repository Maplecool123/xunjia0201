<!--#include file="Inc/Function.asp" -->
<html>
  <head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>时间</title>
  </head>
  <body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" bgcolor=#000000>
    <table border=0 width=100% height=100% cellpadding="0" cellspacing="0">
	  <tr>
	    <td align=center valign=center>
	      <font color=#FFFF00 style='font-size:9pt'>
	        <img alt='系统时间' src=images/time.gif align="texttop">
            <%=date() & " " & hour(now()) & ":" & Minute(now())%></font>
		</td>
	  </tr>
	</table>
	<script>
      setInterval('location.href="main_time.asp"',30000);
	</script>
  </body>
</html>
