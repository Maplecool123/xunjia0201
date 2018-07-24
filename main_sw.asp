<html>
  <head>
    <title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
  function change_menu_button_img(kind)
  {
	document.getElementById('msw_button_img').src=top.menu_sw_button_img[kind].src;
  }
  
  function select_menu_button_img(mode)
  {
    if (top.menu_flag == 1) {
	  if (mode == 1)
	    change_menu_button_img(1);
	  else
	    change_menu_button_img(0);
	}else {
	  if (mode == 1)
	    change_menu_button_img(3);
	  else
	    change_menu_button_img(2);
	}
  }

  function winscreen()
  {
    if (top.menu_flag == 1) {
      top.menu_flag=0;
      top.document.getElementById('screen').cols='0,7,*';
	  change_menu_button_img(2);
    }else {
      top.menu_flag=1;
      top.document.getElementById('screen').cols='150,7,*';
	  change_menu_button_img(0);
    }
  }
</script>
  </head>
  <body onLoad=winscreen(); background='images/sw_bg.png' leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0 bgcolor=#888888 oncontextmenu="window.event.returnValue=false">
    <table width=100% height=100% border=0 cellpadding="0" cellspacing="0">
	  <tr>
	    <td align=center>
		  <img id=msw_button_img border=0 src='' onClick=winscreen(); onMouseOver=select_menu_button_img(1); onMouseOut=select_menu_button_img(0);>
		</td>
	  </tr>
	</table>
  </body>
</html>
