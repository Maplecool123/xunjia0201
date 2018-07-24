<!--#include file="Conn.asp" -->
<!--#include file="Const.asp" -->
<!--#include file="inc/function.asp" -->
<html>
  <head>
    <title><%=sysname%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<script language="javascript">
	  var info_classname='open_data';
	  var clean_info_classname='open_data';
	  var menu_flag=0;
	
	  var menu_sw_button_img = new Array();
	  for (i=0;i<4;i++)
		menu_sw_button_img[i] = new Image;
	  menu_sw_button_img[0].src='images/sw_button1.png';
	  menu_sw_button_img[1].src='images/sw_button2.png';
	  menu_sw_button_img[2].src='images/sw_button3.png';
	  menu_sw_button_img[3].src='images/sw_button4.png';
	  
	  function m_winscreen()
	  {
		mouse_start_addr = (document.body.clientHeight-22)/2-25;
		mouse_end_addr = mouse_start_addr+50;
		mouse_y=event.clientY-22;
		if (menu_flag == 0 && event.clientX == 0 && mouse_y >= mouse_start_addr && mouse_y <= mouse_end_addr)
		  document.getElementById('menu_sw').contentWindow.winscreen();
	  }
	
	  function m_select_menu_button_img(mode)
	  {
		if (mode == 0)
		  document.getElementById('menu_sw').contentWindow.select_menu_button_img(mode);
		else {
		  mouse_start_addr = (document.body.clientHeight-22)/2-25;
		  mouse_end_addr = mouse_start_addr+50;
		  mouse_y=event.clientY-22;
		  if (menu_flag == 0 && event.clientX == 0 && mouse_y >= mouse_start_addr && mouse_y <= mouse_end_addr)
			document.getElementById('menu_sw').contentWindow.select_menu_button_img(mode);
		}
	  }
	</script>
  </head>
<frameset rows="50,*,30" framespacing="0" border="0" frameborder="0" onclick=m_winscreen(); onMouseOver=m_select_menu_button_img(1); onMouseOut=m_select_menu_button_img(0);>
	<frame name="panel" scrolling="no" noresize src="main_panel.asp">
	<frameset id=screen cols="150,7,*">
    		<frameset rows="50,*,22">
        		<frame name="Menu" target="main" src="main_menu_title.asp" scrolling="no">
        		<frame name="MenuFrame" target="main" src="main_menu.asp" scrolling="auto">
        		<frame name="Time" target="main" src="main_time.asp" scrolling="no">
      		</frameset>
      		<frame name="MenuSw" id=menu_sw noresize scrolling="no" frameborder="0" target="MenuFrame" src="main_sw.asp">
      		<frame name="main" id=main src='main_main.asp' scrolling="auto"ght="0">
	</frameset>
	<frame name="foot" id=foot src="main_foot.asp" noresize frameborder="0" scrolling="NO"  marginwidth="0" marginheight="0">
    	<noframes>
      	<body>
        	<p>本网页使用框架，您的浏览器不支持。</p>
      	</body>
    	</noframes>
</frameset>
</html>
<%
call CloseConn()
%>