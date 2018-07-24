<%iflogin=session("iflogin")
session("loginas")=""
session("userrealname")=""
session("username")=""
session("user_id")=""
session("iflogin")=""
session("statue")=""

if iflogin=0 then
	response.redirect("index.asp")
else
	response.redirect("ad_login.asp")
end if
%>