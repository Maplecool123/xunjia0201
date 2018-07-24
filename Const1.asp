<%
'On Error Resume Next
set rs_config=conn.execute("select * from config")
sysname=rs_config("sysname")
sysuser=rs_config("sysuser")
sysmanagercontact=rs_config("sysmanagercontact")
syscontactphone=rs_config("syscontactphone")
sysmanager=rs_config("sysmanager")
gonggao=rs_config("gonggao")
maxrecord=rs_config("maxrecord")
maxrecord1=rs_config("maxrecord1")
maxcfile=rs_config("maxcfile")
maxccontract=rs_config("maxccontract")
maxsubtender=rs_config("maxsubtender")
Modalwidth=rs_config("Modalwidth")
Modalheight=rs_config("Modalheight")
canreg=rs_config("canreg")
usercheck=rs_config("usercheck")
shouldknow=rs_config("shouldknow")
fileprepare=rs_config("fileprepare")
uploadlimit=rs_config("uploadlimit")
rs_config.close
set rs_config=nothing
%>