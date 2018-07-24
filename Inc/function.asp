<%Function Filtersql(Str)
    If Trim(Str) = "" Or IsNull(Str) Then
        Filtersql=""
    Else
    	Str = Replace(Str, "%", "b!f!h")
    	Str = Replace(Str, "#", "b!g!h")
    	Str = Replace(Str, "+", "b!h!h")
    	Str = Replace(Str, "\", "b!i!h")
    	Filtersql = Str
    End If
End Function

Function getsql(Str)
    If Trim(Str) = "" Or IsNull(Str) Then
        getsql=""
    Else
    	Str = Replace(Str, "b!f!h", "%")
    	Str = Replace(Str, "b!g!h", "#")
    	Str = Replace(Str, "b!h!h", "+")
    	Str = Replace(Str, "b!i!h", "\")
    	getsql = Str
    End If
End Function

function sonic(tnum)
if isnumeric(tnum) then
	if instr(tnum,"E")>1 then
		tnum_yuan=tnum
		tnum=formatnumber(tnum,0,,0,0)
		if tnum_yuan-tnum=0 then
			sonic=tnum
		else
			sonic=formatnumber(tnum_yuan,2,,0,0)
		end if
	end if
	
	if tnum>0 then
		if instr(tnum,".")=1 then
			tnum="0"+cstr(tnum)
		end if
		sonic=tnum
	elseif tnum<0 then
		tnum=-tnum
		if instr(tnum,".")=1 then
			tnum="0"+cstr(tnum)
		end if
		sonic="-"&tnum
	else
		sonic=0
	end if
end if
end function

function getmonthstartdate(yy,mm)
if yy<>"" and isnumeric(yy) and mm<>"" and isnumeric(mm) then
	getmonthstartdate=cdate(yy&"-"&mm&"-1")
end if
end function

function getmonthenddate(yy,mm)
if yy<>"" and isnumeric(yy) and mm<>"" and isnumeric(mm) then
	if mm=12 then
		yy=yy+1
		mm=1
	else
		mm=mm+1
	end if
	getmonthenddate=cdate(cdate(yy&"-"&mm&"-1")-1)
	mm=mm-1
end if
end function

sub warnifdate(datestr,titlestr)
if datestr<>"" then
	if isdate(datestr)=false then
		call HintAndBack("您输入的"&titlestr&"时间格式不正确，请重新输入！",1)
	end if
else
	call HintAndBack("您输入的"&titlestr&"时间格式不正确，请重新输入！",1)
end if
end sub

function getcompanystatue(sid)
select case sid
case "1"
getcompanystatue="<font color='#0000ff'>通过审核</font>"
case "2"
getcompanystatue="<font color='#00ff00'>未通过审核</font>"
case "0"
getcompanystatue="<font color='#ff0000'>未完成审核</font>"
case else
getcompanystatue="<font color='#ff0000'>未完成审核</font>"
end select
end function

function getcompanyregstatue(sid)
select case sid
case "1"
getcompanyregstatue="<font color='#ff0000'>已填写基本信息</font>"
case "2"
getcompanyregstatue="<font color='#ff0000'>已上传资质</font>"
case "3"
getcompanyregstatue="<font color='#0000ff'>完成注册</font>"
case else
getcompanyregstatue="<font color='#ff0000'>不明</font>"
end select
end function

function gettenderstatue(sid)
select case sid
case "0"
gettenderstatue="<font color='#0000ff'>尚未开始</font>"
case "1"
gettenderstatue="<font color='#0000ff'>竞价中</font>"
case "2"
gettenderstatue="<font color='#ff0000'>已结标</font>"
case "3"
gettenderstatue="<font color='#0000ff'>流标</font>"
case "4"
gettenderstatue="<font color='#0000ff'>已二次竞价</font>"
case else
gettenderstatue="<font color='#ff0000'>不明</font>"
end select
end function

function getlastdjhno(tenderdate,DataBasetype)
sqlfg="select top 1 djh from tender where id<>0 and ifdel=0"
if DataBasetype=0 then
	sqlfg=sqlfg&" and addtime between #"&tenderdate&"# and #"&tenderdate+1&"#"
else
	sqlfg=sqlfg&" and addtime between '"&tenderdate&"' and '"&tenderdate+1&"'"
end if
sqlfg=sqlfg&" order by djh desc"
set rsfg=conn.execute(sqlfg)
if not rsfg.eof then
	lastdjh=rsfg(0)
	if len(lastdjh)=10 then lastno=cint(right(lastdjh,2))
	getlastdjhno=right("0"&(lastno+1),2)
else
	getlastdjhno="01"
end if
rsfg.close
getlastdjhno=year(tenderdate)&right("0"&month(tenderdate),2)&right("0"&day(tenderdate),2)&getlastdjhno
end function

function checkifnum(no)
if no<>"" and isnumeric(no) then
	checkifnum=no
else
	checkifnum=0
end if
end function

function sishewuru(no,nn)
sishewuru=round(checkifnum(no),nn)
end function

function shownameint(zi,biao,checkzi,id)
if id="" or isnumeric(id)=false then
	shownameint="未知"
else
	sql="select "&zi&" from "&biao&" where "&checkzi&"="&id
	set rss=conn.execute(sql)
	if rss.eof then
		shownameint="未知"
	else
		shownameint=rss(0)
	end if
	rss.close
end if
end function

function shownameintnull(zi,biao,checkzi,id)
if id="" or isnumeric(id)=false then
	shownameintnull=""
else
	sql="select "&zi&" from "&biao&" where "&checkzi&"="&id
	set rss=conn.execute(sql)
	if rss.eof then
		shownameintnull=""
	else
		shownameintnull=rss(0)
	end if
	rss.close
end if
end function

function shownamestr(zi,biao,checkzi,str,fujia)
sql="select "&zi&" from "&biao&" where "&checkzi&"='"&str&"'"&fujia&""
set rss=conn.execute(sql)
if rss.eof then
shownamestr="未知"
else
shownamestr=rss(0)
end if
rss.close
end function

function shownamestrnull(zi,biao,checkzi,str,fujia)
sql="select "&zi&" from "&biao&" where "&checkzi&"='"&str&"'"&fujia&""
set rss=conn.execute(sql)
if rss.eof then
shownamestrnull=""
else
shownamestrnull=rss(0)
end if
rss.close
end function

function showtendergrade(tid)
if tid=0 then
	showtendergrade="所有等级"
else
	set rss=conn.execute("select companygradename from companygrade where id="&tid&"")
	if rss.eof then
		showtendergrade="<font color=red>未知等级</font>"
	else
		showtendergrade=rss(0)
	end if
	rss.close
end if
end function

Public Function PathReplace(Str)
PathReplace=""
if Str<>"" then
	Str=replace(Str,"\/","\")
	'Str=replace(Str,"/\","\")
	Str=replace(Str,"/","\")
	Str=replace(Str,"//","\")
	Str=replace(Str,"\\","\")
	Str=replace(Str,"\\","\")
	PathReplace=Str
end if
End Function

Public Function FileReplace(Str)
FileReplace=""
if Str<>"" then
	Str=replace(Str," ","")
	Str=replace(Str,"　","")
	'Str=replace(Str,",","")
    'Str=replace(Str,"，","")
    'Str=replace(Str,"、","")
    'Str=replace(Str,"。","")
    Str=replace(Str,"~","")
    'Str=replace(Str,"!","")
    Str=replace(Str,"@","")
    Str=replace(Str,"#","")
    Str=replace(Str,"$","")
    Str=replace(Str,"%","")
    Str=replace(Str,"^","")
    Str=replace(Str,"&","")
    Str=replace(Str,"*","")
    'Str=replace(Str,"(","")
    'Str=replace(Str,")","")
    'Str=replace(Str,"-","")
    'Str=replace(Str,"_","")
    'Str=replace(Str,"+","")
    Str=replace(Str,"=","")
    Str=replace(Str,"|","")
    Str=replace(Str,"`","")
    Str=replace(Str,"'","’")
	FileReplace=Str
end if
End Function

Function ScanFile(FilePath)
	ScanFile=False
	dim FSOs,ofile,filetxt,DoMyBest,regEx,Matches,Match
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Function
	if len(filetxt)>0 then
	'特征码检查
	   'Check "WScr"&DoMyBest&"ipt.Shell"
	    If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then ScanFile=True
	   'Check "She"&DoMyBest&"ll.Application"
	    If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then ScanFile=True
		if instr( filetxt,"输入马的内容")>0 or instr(filetxt,"保存文件的<font color=red>绝对路径(包括文件名:如D:\web\x.asp):</font>")>0 then ScanFile=True
		'Check .Encode
		 Set regEx = New RegExp
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
		 If regEx.Test(filetxt) Then ScanFile=True
	    'Check my ASP backdoor :(
		 regEx.Pattern = "\bEv"&"al\b"
		 If regEx.Test(filetxt) Then ScanFile=True
		'Check exe&cute backdoor
		 'regEx.Pattern = "[^.]\bExe"&"cute\b"
		 'If regEx.Test(filetxt) Then ScanFile=True
		 'Set regEx = Nothing
		'Check include file
		 Set regEx = New RegExp
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		 Set Matches = regEx.Execute(filetxt)
		 For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		 Next
		Set Matches = Nothing
		Set regEx = Nothing
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then ScanFile=True
		Set Matches = Nothing
		Set regEx = Nothing
		'Check Crea"&"teObject
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				ScanFile=True
				exit Function
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
	end if
	set ofile = nothing
	set fsos = nothing	
End Function

Public Function CheckFile(strFileSource,sIsDelete)
Set fso = CreateObject("scripting.FileSystemObject") 
If fso.FileExists(strFileSource) Then
	if sIsDelete=True then 
		fso.DeleteFile(strFileSource)
	end if 
End If 
Set fso = Nothing 	
End Function

function getshuliang(zi1,biao,zi2,checkzi,fujia)
set rscb=conn.execute("Select sum("&zi1&") from ["&biao&"] where "&zi2&"="&checkzi&""&fujia&"")
if isnumeric(rscb(0))=false or rscb(0)="" then
	getshuliang=0
else
	getshuliang=rscb(0)
end if
rscb.close
end function

function getshuliang1(zi1,biao,fujia)
set rscb=conn.execute("Select sum("&zi1&") from ["&biao&"] where "&fujia&"")
if isnumeric(rscb(0))=false or rscb(0)="" then
	getshuliang1=0
else
	getshuliang1=checkifnum(rscb(0))
end if
rscb.close
end function

function gettotalitem(biao,zi,checkzi,fujia)
set rscb=conn.execute("Select count(*) from ["&biao&"] where "&zi&"='"&checkzi&"'"&fujia&"")
if isnumeric(rscb(0))=false or rscb(0)="" then
	gettotalitem=0
else
	gettotalitem=rscb(0)
end if
rscb.close
end function

function gettotalitem1(biao,zi,checkzi,fujia)
set rscb=conn.execute("Select count(*) from ["&biao&"] where "&zi&"="&checkzi&""&fujia&"")
if isnumeric(rscb(0))=false or rscb(0)="" then
	gettotalitem1=0
else
	gettotalitem1=rscb(0)
end if
rscb.close
end function

function showadmingrade(ag)
select case ag
case "1"
showadmingrade="计划员"
case "2"
showadmingrade="注册审核员"
case "3"
showadmingrade="纪管审查员"
case "99"
showadmingrade="超级管理员"
case else
showadmingrade="未知"
end select
end  function

function cutstr(str,lenth)
if str="" or isnull(str) then
	if len(str)>lenth then
		cutstr=left(str,lenth)&"…"
	else
		cutstr=str
	end if
else
	cutstr=str
end if
end function

Public Sub getcompanyinfo(cid)
set rs_com=conn.execute("select * from company where id="&cid&"")
if not rs_com.eof then
	companyinfo_grade=rs_com("companygrade")
	companyinfo_class=rs_com("companyclass")
	
	companyinfo_grade_order=0
	set rs_grade=conn.execute("select * from companygrade where id="&companyinfo_grade&"")
	if not rs_grade.eof then
		companyinfo_grade_order=rs_grade("companygradeorder")
	end if
	rs_grade.close
end if
rs_com.close
end Sub

function getmaxcompetitiveprice(danjuhao,fujia)
getmaxcompetitiveprice=0
set rsd=conn.execute("select max(spotprice) from competitive where djh='"&danjuhao&"' and ifzu=1"&fujia&"")
if not rsd.eof then
	getmaxcompetitiveprice=checkifnum(rsd(0))
end if
rsd.close
end function

function getmaxcompetitivecompanyno(danjuhao,fujia)
getmaxcompetitivecompanyno=0
set rsg=conn.execute("select count(*) from competitive where djh='"&danjuhao&"' and ifzu=1"&fujia)
if not rsg.eof then
	getmaxcompetitivecompanyno=checkifnum(rsg(0))
end if
rsg.close
end function

sub getcompanycompetitiveno(cid)   '竞价信息
competitiveno=0
getcompetitiveno=0
getcompetitiveallprice=0
gettime=0
set rsg=conn.execute("select count(*) from competitive where companyid="&cid&" and ifzu=1")
if not rsg.eof then
	competitiveno=checkifnum(rsg(0))
end if
rsg.close

set rsg=conn.execute("select count(*) from competitive where companyid="&cid&" and ifzu=1 and (statue=1 or statue=2)")
if not rsg.eof then
	getcompetitiveno=checkifnum(rsg(0))
end if
rsg.close

set rsg=conn.execute("select sum(spotprice) from competitive where companyid="&cid&" and ifzu=1 and (statue=1 or statue=2)")
if not rsg.eof then
	getcompetitiveallprice=checkifnum(rsg(0))
end if
rsg.close

set rsg=conn.execute("select top 1 gettime from competitive where companyid="&cid&" and (statue=1 or statue=2) and ifzu=1 order by gettime desc")
if not rsg.eof then
	gettime=rsg(0)
end if
rsg.close
end sub

function getcompanycompetitiveinfo(comid,tclass)
set rs_com=conn.execute("select * from company where id="&comid&"")
if not rs_com.eof then
	companyinfo_grade=rs_com("companygrade")
	companyinfo_class=rs_com("companyclass")
	
	companyinfo_grade_order=0
	set rs_grade=conn.execute("select * from companygrade where id="&companyinfo_grade&"")
	if not rs_grade.eof then
		companyinfo_grade_order=rs_grade("companygradeorder")
	end if
	rs_grade.close
end if
rs_com.close

useridlikestr=","&comid&","
sqlcc="select count(*) from tender where ifdel=0 and ifzu=1"
sqlcc1=" and tendergrade<="&companyinfo_grade_order&""
if companyinfo_class<>"综合" then
	sqlcc=sqlcc&" and (tenderclass='"&companyinfo_class&"' or tenderclass='综合')"
end if
select case tclass
case "1"   'recommend
sqlcc=sqlcc&sqlcc1&" and statue=1 and djh not in (select djh from competitive where companyid="&comid&")"
case "2"   'doing
sqlcc=sqlcc&" and statue=1 and djh in (select djh from competitive where companyid="&comid&")"
case "3"   'did
sqlcc=sqlcc&" and (statue=2 or statue=3 or statue=4) and djh in (select djh from competitive where companyid="&comid&" and (statue=0 or statue=3) and ifzu=1)"
case "4"   'get
sqlcc=sqlcc&" and statue=2 and djh in (select djh from competitive where companyid="&comid&" and (statue=1 or statue=2) and ifzu=1)"
case "5"   'future
sqlcc=sqlcc&sqlcc1&" and statue=0"
case "6"   'mycollection
sqlcc=sqlcc&sqlcc1&" and (statue=0 or statue=1) and focusman like '%"&useridlikestr&"%'"
end select
set rsci=conn.execute(sqlcc)
if not rsci.eof then
	getcompanycompetitiveinfo=checkifnum(rsci(0))
else
	getcompanycompetitiveinfo=0
end if

if getcompanycompetitiveinfo>0 then
	getcompanycompetitiveinfo="<font color='#FF0000'>("&getcompanycompetitiveinfo&")</font>"
else
	getcompanycompetitiveinfo="<font color='#0000FF'>("&getcompanycompetitiveinfo&")</font>"
end if
rsci.close
end function

function getallcompetitiveinfo(aclass)
sqlcc="select count(*) from tender where ifzu=1"
select case aclass
case "1"   'future
sqlcc=sqlcc&" and ifdel=0 and statue=0"
case "2"   'doing
sqlcc=sqlcc&" and ifdel=0 and statue=1"
case "3"   'decide
sqlcc=sqlcc&" and ifdel=0 and statue=2 and djh not in (select djh from competitive where (statue=1 or statue=2) and ifzu=1)"
case "4"   'did
sqlcc=sqlcc&" and ifdel=0 and statue=2 and djh in (select djh from competitive where (statue=1 or statue=2) and ifzu=1)"
case "5"   'float
sqlcc=sqlcc&" and ifdel=0 and (statue=3 or statue=4)"
case "6"   'float_del
sqlcc=sqlcc&" and ifdel=1"
end select
set rsci=conn.execute(sqlcc)
if not rsci.eof then
	getallcompetitiveinfo=checkifnum(rsci(0))
else
	getallcompetitiveinfo=0
end if

if getallcompetitiveinfo>0 then
	getallcompetitiveinfo="<font color='#FF0000'>("&getallcompetitiveinfo&")</font>"
else
	getallcompetitiveinfo="<font color='#0000FF'>("&getallcompetitiveinfo&")</font>"
end if
rsci.close
end function

function getfinalcompany(danjuhao)
getfinalcompany=0
set rsd=conn.execute("select companyid from competitive where djh='"&danjuhao&"' and (statue=1 or statue=2) and ifzu=1")
if not rsd.eof then
	getfinalcompany=rsd(0)
end if
rsd.close
end function

function getfinalmoney(danjuhao)
getfinalmoney=0
set rsd=conn.execute("select spotprice from competitive where djh='"&danjuhao&"' and (statue=1 or statue=2) and ifzu=1")
if not rsd.eof then
	getfinalmoney=checkifnum(rsd(0))
end if
rsd.close
end function

Public Sub tendergotorightstatue()
sql_tender="UPDATE tender SET statue=1 WHERE statue=0 and ifdel=0"
if IsSqlDataBase=0 then
	sql_tender=sql_tender&" and ((startdate=#"&date()&"# and startdatehour<="&hour(now())&") or startdate<#"&date()&"#) and ((enddate=#"&date()&"# and enddatehour>"&hour(now())&") or enddate>#"&date()&"#)"
else
	sql_tender=sql_tender&" and ((startdate='"&date()&"' and startdatehour<="&hour(now())&") or startdate<'"&date()&"') and ((enddate='"&date()&"' and enddatehour>"&hour(now())&") or enddate>'"&date()&"')"
end if
conn.execute(sql_tender)

set rst_s=server.createobject("ADODB.RecordSet")
sql_tender="select * from tender where ifdel=0 and (statue=1 or statue=0) and ifzu=1"
if IsSqlDataBase=0 then
	sql_tender=sql_tender&" and ((enddate=#"&date()&"# and enddatehour<="&hour(now())&") or enddate<#"&date()&"#)"
else
	sql_tender=sql_tender&" and ((enddate='"&date()&"' and enddatehour<="&hour(now())&") or enddate<'"&date()&"')"
end if
rst_s.open sql_tender,conn,1,3
if not rst_s.eof then
	do while not rst_s.eof
	if rst_s("minmoney")-getmaxcompetitiveprice(rst_s("djh"),"")<=0 and rst_s("companyno")-getmaxcompetitivecompanyno(rst_s("djh"),"")<=0 then
		'rst_s("statue")=3
		'response.write rst_s("djh")&"-"&getmaxcompetitiveprice(rst_s("djh"))&"--"&getmaxcompetitivecompanyno(rst_s("djh"))&"<br>"
		conn.execute("UPDATE tender SET statue=2 WHERE djh='"&rst_s("djh")&"'")
	else
		'response.write rst_s("djh")&"-"&getmaxcompetitiveprice(rst_s("djh"))&"--"&getmaxcompetitivecompanyno(rst_s("djh"))&"<br>"
		'rst_s("statue")=2
		
		conn.execute("UPDATE tender SET statue=3 WHERE djh='"&rst_s("djh")&"'")
	end if
	rst_s.movenext
	loop
end if
rst_s.close
end Sub

'JS提示后跳转
Sub HintAndTurn(message,gourl)
	Response.Write "<script language=JavaScript>"
	Response.Write "alert(""" & message & """);"
	Response.Write "window.location=""" & gourl & """;"
	Response.Write "</script>"
End Sub

'JS提示后整页跳转
Sub HintAndTurnTopFrame(message,gourl)
	Response.Write "<script language=JavaScript>"
	Response.Write "alert(""" & message & """);"
	Response.Write "window.top.location.href=""" & gourl & """;"
	Response.Write "</script>"
End Sub


'JS提示后关闭窗口
Sub HintAndClose(message)
	Response.Write "<script language=JavaScript>"
	Response.Write "alert(""" & message & """);"
	Response.Write "window.close();"
	Response.Write "</script>"
End Sub


'JS提示后后退
Sub HintAndBack(message,gonum)
	Response.Write "<script language=JavaScript>"
	Response.Write "alert(""" & message & """);"
	Response.Write "history.go(-" & gonum & ");"
	Response.Write "</script>"
	Response.End()
End Sub

'JS提示后不跳转
Sub Hint(message)
	Response.Write "<script language=JavaScript>alert(""" & message & """);</script>"
End Sub


'过滤：,(){}`[]-+*%/="'~!&|\<>?:;.@#$^空格
Function ReplaceBadStr(chkstr)
	If chkstr = "" or isnull(chkstr) Then
		ReplaceBadStr = ""
	Else
		ReplaceBadStr = Replace(chkstr,"{","")
		ReplaceBadStr = Replace(ReplaceBadStr,"}","")
		ReplaceBadStr = Replace(ReplaceBadStr,"`","")
		ReplaceBadStr = Replace(ReplaceBadStr,"[","")
		ReplaceBadStr = Replace(ReplaceBadStr,"]","")
		ReplaceBadStr = Replace(ReplaceBadStr,",","")
		ReplaceBadStr = Replace(ReplaceBadStr,"(","")
		ReplaceBadStr = Replace(ReplaceBadStr,")","")
		'ReplaceBadStr = Replace(ReplaceBadStr,"-","")
		ReplaceBadStr = Replace(ReplaceBadStr,"+","")
		ReplaceBadStr = Replace(ReplaceBadStr,"*","")
		ReplaceBadStr = Replace(ReplaceBadStr,"%","")
		ReplaceBadStr = Replace(ReplaceBadStr,"/","")
		ReplaceBadStr = Replace(ReplaceBadStr,"=","")
		ReplaceBadStr = Replace(ReplaceBadStr,chr(34),"")
		ReplaceBadStr = Replace(ReplaceBadStr,"'","")
		ReplaceBadStr = Replace(ReplaceBadStr,"~","")
		ReplaceBadStr = Replace(ReplaceBadStr,"!","")
		ReplaceBadStr = Replace(ReplaceBadStr,"&","")
		ReplaceBadStr = Replace(ReplaceBadStr,"|","")
		ReplaceBadStr = Replace(ReplaceBadStr,"\","")
		ReplaceBadStr = Replace(ReplaceBadStr,"<","")
		ReplaceBadStr = Replace(ReplaceBadStr,">","")
		ReplaceBadStr = Replace(ReplaceBadStr,"?","")
		ReplaceBadStr = Replace(ReplaceBadStr,":","")
		ReplaceBadStr = Replace(ReplaceBadStr,";","")
		ReplaceBadStr = Replace(ReplaceBadStr,".","")
		ReplaceBadStr = Replace(ReplaceBadStr,"@","")
		ReplaceBadStr = Replace(ReplaceBadStr,"#","")
		ReplaceBadStr = Replace(ReplaceBadStr,"$","")
		ReplaceBadStr = Replace(ReplaceBadStr,"^","")
		'ReplaceBadStr = Replace(ReplaceBadStr," ","")
	End If
End Function


'仿SQL的IIf函数
Function iIf(bis,truevalue,falsevalue)
	If bis Then iIf = truevalue Else iIf = falsevalue
End Function

'判断是否为合法的字符串：abcdefghijklmnopqrstuvwxyz0123456789_
Function StrValid(str)
	Dim i,c
	StrValid=True
	If str="" Then
		StrValid=False
	Else
		For i = 1 to Len(str)
			c = Lcase(Mid(str, i, 1))
			If InStr("abcdefghijklmnopqrstuvwxyz0123456789_",c) <= 0 Then StrValid=False
			Exit For
		Next
	End If
End Function

'判断字符串长度，如果超过规定长度，截短并加省略号。
Function ShortString(szStr,nLength)
	If szStr = "" Then
		ShortString = ""
	Else
		If Len(szStr) > nLength Then
			ShortString = Left(szStr,nLength-1)&".."
		Else
			ShortString = szStr
		End If
	End If
End Function

'计算字符长度
Function StrLength(Str,chinese)
	If IsNull(Str & "") Or Trim(Str & "")="" Then StrLength=0
	If chinese Then
		StrLength=Len(str)
	Else 
		Dim l,t,c
		Dim i
		l = Len(str & "")
		t = l
		For i=1 To l
			C=Asc(Mid(Str,i,1))
			If c<0 Then c=c+65536
			If c>255 Then
				t=t+1
			End If
		Next
		StrLength=t
	End If
End Function

'将内容保存为文件
Function StringToFile(WriteString,FileName)
	If WriteString = "" Or FileName = "" Then Exit Function
	Dim ffso,wfso
	Set ffso = server.createobject("scripting.filesystemobject")
	Set wfso = ffso.CreateTextFile(server.mappath(FileName))
	wfso.Writeline(WriteString)
	wfso.Close
	Set wfso = Nothing
	Set ffso = Nothing
End Function 

'计算两个时间差
Function DiffDate(StartDate,EndDate,DateType)
	'格式化为(StartDate-EndDate)个DateType单位
	'DateType单位为：yyyy---年;q---季度;n---月;y---一年的日数;d---日;w---一周的日数;ww---周;h---小时;m---分钟;s---秒
	DiffDate = DateDiff(DateType,EndDate,StartDate)
End Function

'取地址栏值
Function GetUrl(action,isWrite)
	Select Case action
		Case 1			'赋“http://”+域名
			GetUrl = "http://" & Request.servervariables("HTTP_HOST")
			If isWrite Then Response.Write GetUrl
		Case 2			'赋文件名
			GetUrl = Mid(Request.ServerVariables("script_name"),InstrRev(Replace(Request.ServerVariables("script_name"),"\","/"),"/")+1)
			If isWrite Then Response.Write GetUrl
		Case 3			'赋地址栏参数
			GetUrl = Request.servervariables("QUERY_STRING")
			If isWrite Then Response.Write GetUrl
		Case 4			'赋文件名+地址栏参数
			GetUrl = Mid(Request.ServerVariables("script_name"),InstrRev(Replace(Request.ServerVariables("script_name"),"\","/"),"/")+1)
			If Request.servervariables("QUERY_STRING")<>"" Then GetUrl = GetUrl & "?" & Request.servervariables("QUERY_STRING")
			If isWrite Then Response.Write GetUrl
		Case 5			'赋目录及文件名
			GetUrl = Request.servervariables("script_name")
			If isWrite Then Response.Write GetUrl
		Case 6			'赋“http://”+域名+目录及文件名
			GetUrl = "http://" & Request.servervariables("HTTP_HOST") & Request.servervariables("script_name")
			If isWrite Then Response.Write GetUrl
		Case Else		'赋整个地址栏值
			GetUrl = "http://" & Request.servervariables("HTTP_HOST") & Request.servervariables("script_name")
			If Request.servervariables("QUERY_STRING")<>"" Then GetUrl = GetUrl & "?" & Request.servervariables("QUERY_STRING")
			If isWrite Then Response.Write GetUrl
	End Select 
End Function

'保留html的<img>和<br>标签，其它全部过滤
Function HoldHTML(strHTML)
	if strHTML ="" or isNull(strHTML) then exit Function
	Dim objRegExp,Match,Matches
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<(?!img|br).+?>"
	Set Matches = objRegExp.Execute(strHTML)
		For Each Match In Matches
			strHtml = Replace(strHTML,Match.Value,"")
		Next
	HoldHTML = strHTML
	Set objRegExp = Nothing
End Function

'过滤所有html标签
Function RemoveHTML(strHTML)
	Dim objRegExp,Match,Matches
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<.+?>"
	Set Matches = objRegExp.Execute(strHTML)
		For Each Match In Matches
			strHtml = Replace(strHTML,Match.Value,"")
		Next
	RemoveHTML = strHTML
	Set objRegExp = Nothing
End Function


'取上页来源
Function PrevUrl(isAll,isWrite)
	PrevUrl = Trim(Request("PrevUrl"))
	If PrevUrl <> "" Then
		'值未变
	ElseIf Request.ServerVariables("HTTP_REFERER") <> "" Then
		If isAll Then
			PrevUrl = Request.ServerVariables("HTTP_REFERER")
		Else
			PrevUrl = Split(Request.ServerVariables("HTTP_REFERER"),"/")(UBound(Split(Request.ServerVariables("HTTP_REFERER"),"/")))
		End If
	Else
		PrevUrl = "index.asp"
	End If
	If isWrite Then Response.Write PrevUrl
End Function


'是否外部提交表单
Function ChkPost()
	Dim server_v1,server_v2
	Chkpost = False 
	server_v1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2 = Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(server_v1,8,len(server_v2)) = server_v2 Then Chkpost=True 
End Function


'将非法字符替换为星号
function debadstr(str)
	dim rsdeb,badstr,i
	debadstr=str
	set rsdeb=conn.execute("select blogFiltrate from Blog_Filtrate")
	badstr=split(rsdeb(0),"|")
	for i=0 to ubound(badstr)
	debadstr=replace(debadstr,badstr(i),"***含非法字符***")
	next
	set rsdeb=nothing
end function


'提示
Function MsgBox(title,about,message,whereisgo,gourl,gotitle,isEnd)
	Dim Ci
	message = Split(message,"|")
	MsgBox = vbcrlf & "<center><div style=""text-align:center;border:1px solid #d1d0d0;width:772px;background:#ffffff;"">"
	MsgBox = MsgBox & vbcrlf & "	<div style=""background:#f4f4f4;color:#000000;font-size: 14px;height:24px;font-weight:bold;padding:5px 0px 0px 0px;margin:0px;"">" & title & "</div>"
	MsgBox = MsgBox & vbcrlf & "	<div style=""font-size: 12px;font-weight:normal;text-align:left;background:#ffffff;padding:6px;"">"
	MsgBox = MsgBox & vbcrlf & "		" & about & "<br />"
	MsgBox = MsgBox & vbcrlf & "		<ul>"
	For Ci=0 To Ubound(message)
		If message(Ci)<>"" Then
			MsgBox = MsgBox & vbcrlf & "			<li>" & message(Ci) & "</li>"
		End If
	Next
	MsgBox = MsgBox & vbcrlf & "		</ul>"
	Select Case whereisgo
		Case "gourl"
			MsgBox = MsgBox & vbcrlf & "		<center><a href=""" & gourl & """><span class=""fontwebdings"">q</span>" & gotitle & "</a></center>"
		Case "close"
			MsgBox = MsgBox & vbcrlf & "		<center><a href=""javascript:window.close();""><span style=""font-family:'Webdings';"">r</span>关闭</a></center>"
		Case Else
			MsgBox = MsgBox & vbcrlf & "		<center><a href=""javascript:history.go(-1);""><span style=""font-family:'Webdings';"">7</span>返回</a></center>"
	End Select
	MsgBox = MsgBox & vbcrlf & "	</div>"
	MsgBox = MsgBox & vbcrlf & "</div></center>" & vbcrlf
	Response.Write MsgBox
	If isEnd Then Response.End()
End Function

'过滤非法HTML代码及注入关键字
Function ReplaceBadHtml(str)
		If Isnull(Str) Then
			ReplaceBadHtml=""
			Exit Function 
		End If
		str = replace(str,chr(0),"", 1, -1, 1)
		str = replace(str, """", "&quot;", 1, -1, 1)
		str = replace(str,"<","&lt;", 1, -1, 1)
		str = replace(str,">","&gt;", 1, -1, 1)
		str = replace(str, "script", "&#115;cript", 1, -1, 0)
		str = replace(str, "script", "&#083;cript", 1, -1, 0)
		str = replace(str, "script", "&#083;cript", 1, -1, 0)
		str = replace(str, "script", "&#083;cript", 1, -1, 1)
		str = replace(str, "object", "&#111;bject", 1, -1, 0)
		str = replace(str, "object", "&#079;bject", 1, -1, 0)
		str = replace(str, "object", "&#079;bject", 1, -1, 0)
		str = replace(str, "object", "&#079;bject", 1, -1, 1)
		str = replace(str, "applet", "&#097;pplet", 1, -1, 0)
		str = replace(str, "applet", "&#065;pplet", 1, -1, 0)
		str = replace(str, "applet", "&#065;pplet", 1, -1, 0)
		str = replace(str, "applet", "&#065;pplet", 1, -1, 1)
		str = replace(str, "[", "&#091;")
		str = replace(str, "]", "&#093;")
		str = replace(str, """", "", 1, -1, 1)
		str = replace(str, "=", "&#061;", 1, -1, 1)
		str = replace(str, "", "", 1, -1, 1)
		str = replace(str, "select", "sel&#101;ct", 1, -1, 1)
		str = replace(str, "execute", "&#101xecute", 1, -1, 1)
		str = replace(str, "exec", "&#101xec", 1, -1, 1)
		str = replace(str, "join", "jo&#105;n", 1, -1, 1)
		str = replace(str, "union", "un&#105;on", 1, -1, 1)
		str = replace(str, "where", "wh&#101;re", 1, -1, 1)
		str = replace(str, "insert", "ins&#101;rt", 1, -1, 1)
		str = replace(str, "delete", "del&#101;te", 1, -1, 1)
		str = replace(str, "update", "up&#100;ate", 1, -1, 1)
		str = replace(str, "like", "lik&#101;", 1, -1, 1)
		str = replace(str, "drop", "dro&#112;", 1, -1, 1)
		str = replace(str, "create", "cr&#101;ate", 1, -1, 1)
		str = replace(str, "rename", "ren&#097;me", 1, -1, 1)
		str = replace(str, "count", "co&#117;nt", 1, -1, 1)
		str = replace(str, "chr", "c&#104;r", 1, -1, 1)
		str = replace(str, "mid", "m&#105;d", 1, -1, 1)
		str = replace(str, "truncate", "trunc&#097;te", 1, -1, 1)
		str = replace(str, "nchar", "nch&#097;r", 1, -1, 1)
		str = replace(str, "char", "ch&#097;r", 1, -1, 1)
		str = replace(str, "alter", "alt&#101;r", 1, -1, 1)
		str = replace(str, "cast", "ca&#115;t", 1, -1, 1)
		str = replace(str, "exists", "e&#120;ists", 1, -1, 1)
		str = replace(str,chr(13),"<br>", 1, -1, 1)
		ReplaceBadHtml = replace(str,"'","", 1, -1, 1)
End Function

'检测是否只包含英文和数字
Function IsValidChars(str)
	Dim re,chkstr
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="[^_\.a-zA-Z\d]"
	IsValidChars=True
	chkstr=re.Replace(str,"")
	if chkstr<>str then IsValidChars=False
	set re=nothing
End Function

'检测是否有效的数字
Function IsInteger(Para) 
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

'加亮关键字
Function highlight(byVal strContent,byRef arrayWords)
	Dim intCounter,strTemp,intPos,intTagLength,intKeyWordLength,bUpdate
	if len(arrayWords)<1 then highlight=strContent:exit function
	For intPos = 1 to Len(strContent)
		bUpdate = False
		If Mid(strContent, intPos, 1) = "<" Then
			On Error Resume Next
			intTagLength = (InStr(intPos, strContent, ">", 1) - intPos)
			if err then
			  highlight=strContent
			  err.clear
			end if
			strTemp = strTemp & Mid(strContent, intPos, intTagLength)
			intPos = intPos + intTagLength
		End If
			If arrayWords <> "" Then
				intKeyWordLength = Len(arrayWords)
				If LCase(Mid(strContent, intPos, intKeyWordLength)) = LCase(arrayWords) Then
					strTemp = strTemp & "<span class=""high1"">" & Mid(strContent, intPos, intKeyWordLength) & "</span>"
					intPos = intPos + intKeyWordLength - 1
					bUpdate = True
				End If
			End If
		If bUpdate = False Then
			strTemp = strTemp & Mid(strContent, intPos, 1)
		End If
	Next
	highlight = strTemp
End Function

'过滤超链接
Function checkURL(ByVal ChkStr)
	Dim str:str=ChkStr
	str=Trim(str)
	If IsNull(str) Then
		checkURL = ""
		Exit Function 
	End If
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(d)(ocument\.cookie)"
    Str = re.replace(Str,"$1ocument cookie")
	re.Pattern="(d)(ocument\.write)"
    Str = re.replace(Str,"$1ocument write")
   	re.Pattern="(s)(cript:)"
    Str = re.replace(Str,"$1cri&#112;t ")
   	re.Pattern="(s)(cript)"
    Str = re.replace(Str,"$1cri&#112;t")
   	re.Pattern="(o)(bject)"
    Str = re.replace(Str,"$1bj&#101;ct")
   	re.Pattern="(a)(pplet)"
    Str = re.replace(Str,"$1ppl&#101;t")
   	re.Pattern="(e)(mbed)"
    Str = re.replace(Str,"$1mb&#101;d")
	Set re=Nothing
   	Str = Replace(Str, ">", "&gt;")
	Str = Replace(Str, "<", "&lt;")
	checkURL=Str    
end Function

'过滤特殊字符
Function OnCheckStr(byVal ChkStr) 
	Dim Str:Str=ChkStr
	If IsNull(Str) Then
		OnCheckStr = ""
		Exit Function 
	End If
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str,"'","&#39;")
    Str = Replace(Str,"""","&#34;")
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(w)(here)"
    Str = re.replace(Str,"$1h&#101;re")
	re.Pattern="(s)(elect)"
    Str = re.replace(Str,"$1el&#101;ct")
	re.Pattern="(i)(nsert)"
    Str = re.replace(Str,"$1ns&#101;rt")
	re.Pattern="(c)(reate)"
    Str = re.replace(Str,"$1r&#101;ate")
	re.Pattern="(d)(rop)"
    Str = re.replace(Str,"$1ro&#112;")
	re.Pattern="(a)(lter)"
    Str = re.replace(Str,"$1lt&#101;r")
	re.Pattern="(d)(elete)"
    Str = re.replace(Str,"$1el&#101;te")
	re.Pattern="(u)(pdate)"
    Str = re.replace(Str,"$1p&#100;ate")
	re.Pattern="(\s)(or)"
    Str = re.replace(Str,"$1o&#114;")
	Set re=Nothing
	OnCheckStr=Str
End Function

'恢复特殊字符
Function UnCheckStr(ByVal Str)
		If IsNull(Str) Then
			UnCheckStr = ""
			Exit Function 
		End If
	    Str = Replace(Str,"&#39;","'")
        Str = Replace(Str,"&#34;","""")
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(w)(h&#101;re)"
	    str = re.replace(str,"$1here")
		re.Pattern="(s)(el&#101;ct)"
	    str = re.replace(str,"$1elect")
		re.Pattern="(i)(ns&#101;rt)"
	    str = re.replace(str,"$1nsert")
		re.Pattern="(c)(r&#101;ate)"
	    str = re.replace(str,"$1reate")
		re.Pattern="(d)(ro&#112;)"
	    str = re.replace(str,"$1rop")
		re.Pattern="(a)(lt&#101;r)"
	    str = re.replace(str,"$1lter")
		re.Pattern="(d)(el&#101;te)"
	    str = re.replace(str,"$1elete")
		re.Pattern="(u)(p&#100;ate)"
	    str = re.replace(str,"$1pdate")
		re.Pattern="(\s)(o&#114;)"
	    Str = re.replace(Str,"$1or")
		Set re=Nothing
        Str = Replace(Str, "&amp;", "&")
    	UnCheckStr=Str
End Function

'转换HTML代码
Function HTMLEncode(ByVal reString) 
	Dim Str:Str=reString
	If Not IsNull(Str) Then
   		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
	    Str = Replace(Str, CHR(9), "&#160;&#160;&#160;&#160;")
	    Str = Replace(Str, CHR(32), "&nbsp;")
	    Str = Replace(Str, CHR(39), "&#39;")
    	Str = Replace(Str, CHR(34), "&quot;")
		Str = Replace(Str, CHR(13), "")
		Str = Replace(Str, CHR(10), "<br/>")
		HTMLEncode = Str
	End If
End Function

'反转换HTML代码
Function HTMLDecode(ByVal reString) 
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = Replace(Str, "&gt;", ">")
		Str = Replace(Str, "&lt;", "<")
		Str = Replace(Str, "&#160;&#160;&#160;&#160;", CHR(9))
	    Str = Replace(Str, "&nbsp;", CHR(32))
		Str = Replace(Str, "&#39;", CHR(39))
		Str = Replace(Str, "&quot;", CHR(34))
		Str = Replace(Str, "", CHR(13))
		Str = Replace(Str, "<br/>", CHR(10))
		HTMLDecode = Str
	End If
End Function

'日期转换函数
'ShowType参数：Y-m-d | Y-m-d H:I A | Y-m-d H:I:S | YmdHIS | ym | d | ymd | mdy | w,d m y H:I:S | y-m-dTH:I:S
Function DateToStr(DateTime,ShowType)  
	Dim DateMonth,DateDay,DateHour,DateMinute,DateWeek,DateSecond
	Dim FullWeekday,shortWeekday,Fullmonth,Shortmonth,TimeZone1,TimeZone2
	TimeZone1="+0800"
	TimeZone2="+08:00"
	FullWeekday=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
	shortWeekday=Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
    Fullmonth=Array("January","February","March","April","May","June","July","August","September","October","November","December")
    Shortmonth=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

	DateMonth=Month(DateTime)
	DateDay=Day(DateTime)
	DateHour=Hour(DateTime)
	DateMinute=Minute(DateTime)
	DateWeek=weekday(DateTime)
	DateSecond=Second(DateTime)
	If Len(DateMonth)<2 Then DateMonth="0"&DateMonth
	If Len(DateDay)<2 Then DateDay="0"&DateDay
	If Len(DateMinute)<2 Then DateMinute="0"&DateMinute
	Select Case ShowType
	Case "Y-m-d"  
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay
	Case "Y-m-d H:I A"
		Dim DateAMPM
		If DateHour>12 Then 
			DateHour=DateHour-12
			DateAMPM="PM"
		Else
			DateHour=DateHour
			DateAMPM="AM"
		End If
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
	Case "Y-m-d H:I:S"
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
	Case "YmdHIS"
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond	
	Case "ym"
		DateToStr=Right(Year(DateTime),2)&DateMonth
	Case "d"
		DateToStr=DateDay
    Case "ymd"
        DateToStr=Right(Year(DateTime),4)&DateMonth&DateDay
    Case "mdy" 
        Dim DayEnd
        select Case DateDay
         Case 1 
          DayEnd="st"
         Case 2
          DayEnd="nd"
         Case 3
          DayEnd="rd"
         Case Else
          DayEnd="th"
        End Select 
        DateToStr=Fullmonth(DateMonth-1)&" "&DateDay&DayEnd&" "&Right(Year(DateTime),4)
    Case "w,d m y H:I:S" 
		DateSecond=Second(DateTime)
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
        DateToStr=shortWeekday(DateWeek-1)&","&DateDay&" "& Left(Fullmonth(DateMonth-1),3) &" "&Right(Year(DateTime),4)&" "&DateHour&":"&DateMinute&":"&DateSecond&" "&TimeZone1
    Case "y-m-dTH:I:S"
		If Len(DateHour)<2 Then DateHour="0"&DateHour	
		If Len(DateSecond)<2 Then DateSecond="0"&DateSecond
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&TimeZone2
	Case Else
		If Len(DateHour)<2 Then DateHour="0"&DateHour
		DateToStr=Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
	End Select
End Function

'切割内容 - 按行分割
Function SplitLines(byVal Content,byVal ContentNums) 
	Dim ts,i,l
	ContentNums=int(ContentNums)
	If IsNull(Content) Then Exit Function
	i=1
	ts = 0
	For i=1 to Len(Content)
      l=Lcase(Mid(Content,i,5))
      	If l="<br/>" Then
         	ts=ts+1
      	End If
      l=Lcase(Mid(Content,i,4))
      	If l="<br>" Then
         	ts=ts+1
      	End If
      l=Lcase(Mid(Content,i,3))
      	If l="<p>" Then
         	ts=ts+1
      	End If
	If ts>ContentNums Then Exit For 
	Next
	If ts>ContentNums Then
    	Content=Left(Content,i-1)
	End If
	SplitLines=Content
End Function

'切割内容 - 按字符分割
Function CutStr(byVal Str,byVal StrLen)
	Dim l,t,c,i
	If IsNull(Str) Then CutStr="":Exit Function
	l=Len(str)
	StrLen=int(StrLen)
	t=0
	For i=1 To l
		c=Asc(Mid(str,i,1))
		If c<0 Or c>255 Then t=t+2 Else t=t+1
		IF t>=StrLen Then
			CutStr=left(Str,i)&"..."
			Exit For
		Else
			CutStr=Str
		End If
	Next
End Function

'获取客户端浏览器信息
function getBrowser(strUA) 
 dim arrInfo,strType,temp1,temp2
 strType=""
 strUA=LCase(strUA)
 arrInfo=Array("Unkown","Unkown")
 '浏览器判断
    if Instr(strUA,"mozilla")>0 then arrInfo(0)="Mozilla"
    if Instr(strUA,"icab")>0 then arrInfo(0)="iCab"
    if Instr(strUA,"lynx")>0 then arrInfo(0)="Lynx"
    if Instr(strUA,"links")>0 then arrInfo(0)="Links"
    if Instr(strUA,"elinks")>0 then arrInfo(0)="ELinks"
    if Instr(strUA,"jbrowser")>0 then arrInfo(0)="JBrowser"
    if Instr(strUA,"konqueror")>0 then arrInfo(0)="konqueror"
    if Instr(strUA,"wget")>0 then arrInfo(0)="wget"
    if Instr(strUA,"ask jeeves")>0 or Instr(strUA,"teoma")>0 then arrInfo(0)="Ask Jeeves/Teoma"
    if Instr(strUA,"wget")>0 then arrInfo(0)="wget"
    if Instr(strUA,"opera")>0 then arrInfo(0)="opera"

    if Instr(strUA,"gecko")>0 then 
      strType="[Gecko]"
      arrInfo(0)="Mozilla"
      if Instr(strUA,"aol")>0 then arrInfo(0)="AOL"
      if Instr(strUA,"netscape")>0 then arrInfo(0)="Netscape"
      if Instr(strUA,"firefox")>0 then arrInfo(0)="FireFox"
      if Instr(strUA,"chimera")>0 then arrInfo(0)="Chimera"
      if Instr(strUA,"camino")>0 then arrInfo(0)="Camino"
      if Instr(strUA,"galeon")>0 then arrInfo(0)="Galeon"
      if Instr(strUA,"k-meleon")>0 then arrInfo(0)="K-Meleon"
      arrInfo(0)=arrInfo(0)+strType
   end if
   
   if Instr(strUA,"bot")>0 or Instr(strUA,"crawl")>0 then 
      strType="[Bot/Crawler]"
      arrInfo(0)=""
      if Instr(strUA,"grub")>0 then arrInfo(0)="Grub"
      if Instr(strUA,"googlebot")>0 then arrInfo(0)="GoogleBot"
      if Instr(strUA,"msnbot")>0 then arrInfo(0)="MSN Bot"
      if Instr(strUA,"slurp")>0 then arrInfo(0)="Yahoo! Slurp"
      arrInfo(0)=arrInfo(0)+strType
  end if
  
  if Instr(strUA,"applewebkit")>0 then 
      strType="[AppleWebKit]"
      arrInfo(0)=""
      if Instr(strUA,"omniweb")>0 then arrInfo(0)="OmniWeb"
      if Instr(strUA,"safari")>0 then arrInfo(0)="Safari"
      arrInfo(0)=arrInfo(0)+strType
  end if 
  
  if Instr(strUA,"msie")>0 then 
      strType="[MSIE"
      temp1=mid(strUA,(Instr(strUA,"msie")+4),6)
      temp2=Instr(temp1,";")
      temp1=left(temp1,temp2-1)
      strType=strType & temp1 &"]"
      arrInfo(0)="Internet Explorer"
      if Instr(strUA,"msn")>0 then arrInfo(0)="MSN"
      if Instr(strUA,"aol")>0 then arrInfo(0)="AOL"
      if Instr(strUA,"webtv")>0 then arrInfo(0)="WebTV"
      if Instr(strUA,"myie2")>0 then arrInfo(0)="MyIE2"
      if Instr(strUA,"maxthon")>0 then arrInfo(0)="Maxthon"
      if Instr(strUA,"gosurf")>0 then arrInfo(0)="GoSurf"
      if Instr(strUA,"netcaptor")>0 then arrInfo(0)="NetCaptor"
      if Instr(strUA,"sleipnir")>0 then arrInfo(0)="Sleipnir"
      if Instr(strUA,"avant browser")>0 then arrInfo(0)="AvantBrowser"
      if Instr(strUA,"greenbrowser")>0 then arrInfo(0)="GreenBrowser"
      if Instr(strUA,"slimbrowser")>0 then arrInfo(0)="SlimBrowser"
      arrInfo(0)=arrInfo(0)+strType
   end if
 
 '操作系统判断
    if Instr(strUA,"windows")>0 then arrInfo(1)="Windows"
    if Instr(strUA,"windows ce")>0 then arrInfo(1)="Windows CE"
    if Instr(strUA,"windows 95")>0 then arrInfo(1)="Windows 95"
    if Instr(strUA,"win98")>0 then arrInfo(1)="Windows 98"
    if Instr(strUA,"windows 98")>0 then arrInfo(1)="Windows 98"
    if Instr(strUA,"windows 2000")>0 then arrInfo(1)="Windows 2000"
    if Instr(strUA,"windows xp")>0 then arrInfo(1)="Windows XP"

    if Instr(strUA,"windows nt")>0 then
      arrInfo(1)="Windows NT"
      if Instr(strUA,"windows nt 5.0")>0 then arrInfo(1)="Windows 2000"
      if Instr(strUA,"windows nt 5.1")>0 then arrInfo(1)="Windows XP"
      if Instr(strUA,"windows nt 5.2")>0 then arrInfo(1)="Windows 2003"
    end if
    if Instr(strUA,"x11")>0 or Instr(strUA,"unix")>0 then arrInfo(1)="Unix"
    if Instr(strUA,"sunos")>0 or Instr(strUA,"sun os")>0 then arrInfo(1)="SUN OS"
    if Instr(strUA,"powerpc")>0 or Instr(strUA,"ppc")>0 then arrInfo(1)="PowerPC"
    if Instr(strUA,"macintosh")>0 then arrInfo(1)="Mac"
    if Instr(strUA,"mac osx")>0 then arrInfo(1)="MacOSX"
    if Instr(strUA,"freebsd")>0 then arrInfo(1)="FreeBSD"
    if Instr(strUA,"linux")>0 then arrInfo(1)="Linux"
    if Instr(strUA,"palmsource")>0 or Instr(strUA,"palmos")>0 then arrInfo(1)="PalmOS"
    if Instr(strUA,"wap ")>0 then arrInfo(1)="WAP"
  
 getBrowser=arrInfo
end Function

'计算随机数
function randomStr(intLength)
    dim strSeed,seedLength,pos,str,i
    strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
    seedLength=len(strSeed)
    str=""
    Randomize
    for i=1 to intLength
     str=str+mid(strSeed,int(seedLength*rnd)+1,1)
    next
    randomStr=str
end Function

'检测系统组件是否可用
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
	If IsObjInstalled Then
		IsObjInstalled = true
	Else
		IsObjInstalled = false
	End If
End Function

'广告条是Flash还是图片
function FlashOrPic(Url,DirUrl,FileWidth,FileHeight)
	select case Lcase(Right(Url,3))
		case "jpg","gif","bmp","png"
			if Lcase(left(Url,6)) = "http:/" or Lcase(left(Url,6)) = "ftp://" then
				FlashOrPic = "<img src=""" & Url & """ width=""" & FileWidth & """ height=""" & FileHeight & """ border=""0"">"
			else
				FlashOrPic = "<img src=""" & DirUrl & Url & """ width=""" & FileWidth & """ height=""" & FileHeight & """ border=""0"">"
			end if
		case "swf"
			if Lcase(left(Url,6)) = "http:/" or Lcase(left(Url,6)) = "ftp://" then
				FlashOrPic = "		<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""" & FileWidth & """ height=""" & FileHeight & """>" & vbcrlf & "		  <param name=""movie"" value=""" & Url & """ />" & vbcrlf & "		  <param name=""quality"" value=""high"" />" & vbcrlf & "		  <embed src=""" & Url & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""" & FileWidth & """ height=""" & FileHeight & """></embed>" & vbcrlf & "		</object>" & vbcrlf
			else
				FlashOrPic = "		<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""" & FileWidth & """ height=""" & FileHeight & """>" & vbcrlf & "		  <param name=""movie"" value=""" & DirUrl & Url & """ />" & vbcrlf & "		  <param name=""quality"" value=""high"" />" & vbcrlf & "		  <embed src=""" & DirUrl & Url & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""" & FileWidth & """ height=""" & FileHeight & """></embed>" & vbcrlf & "		</object>" & vbcrlf
			end if
		case else
			FlashOrPic = "错误格式！"
	end select
end function

'加密BT---变态，HOHO……
Function PassWordToBT(pswstr)
	Dim Ci,Cj,PassStr,PwsAsc,PswAscToStr,PswAscStr,AnsiStr
	PassWordToBT = ""
	pswstr = Replace(pswstr," ","")
	If trim(pswstr) <> "" Then
		For Ci = 1 To Len(PswStr)
			PassStr = ""
			PswAscStr = Cstr(Asc(Lcase(Mid(PswStr,Ci,1))))
			For Cj = 1 To Len(PswAscStr)
				PswAscToStr=Mid(PswAscStr,Cj,1)
				Select Case PswAscToStr
				   Case "1":AnsiStr = "マ"
				   Case "2":AnsiStr = "ミ"
				   Case "3":AnsiStr = "ョ"
				   Case "4":AnsiStr = "メ"
				   Case "5":AnsiStr = "モ"
				   Case "6":AnsiStr = "ャ"
				   Case "7":AnsiStr = "ュ"
				   Case "8":AnsiStr = "ム"
				   Case "9":AnsiStr = "ヮ"
				   Case "0":AnsiStr = "ヲ"
				End Select
				PassStr = PassStr & AnsiStr
			 Next
			 If PassWordToBT = "" Then
				PassWordToBT = PassStr
			 Else
				PassWordToBT = PassWordToBT & "ヰ" & PassStr
			 End If
		Next
	Else
		Response.Write "加密内容为空，加密失败！（操作中断！）"
		Response.End()
	End If
End Function

'解密BT---变态，HOHO……
Function PassWordFromBT(pswstr)
	Dim Ci,Cj,PassStr,PasAsc,PswAscToStr,PswAscStr,AnsiStr
	PassWordFromBT = ""
	If trim(pswstr) <> "" Then
		PasAsc = split(pswstr,"ヰ")
		For Ci = 0 To Ubound(PasAsc)
			PassStr = ""
			PswAscStr = PasAsc(Ci)
			For Cj = 1 To Len(PswAscStr)
				PswAscToStr = Mid(PswAscStr,Cj,1)
				Select Case PswAscToStr
				  Case "マ":AnsiStr = "1"
				  Case "ミ":AnsiStr = "2"
				  Case "ョ":AnsiStr = "3"
				  Case "メ":AnsiStr = "4"
				  Case "モ":AnsiStr = "5"
				  Case "ャ":AnsiStr = "6"
				  Case "ュ":AnsiStr = "7"
				  Case "ム":AnsiStr = "8"
				  Case "ヮ":AnsiStr = "9"
				  Case "ヲ":AnsiStr = "0"
				  Case Else:
					Response.Write "解密内容含非法字符，解密失败！（操作中断！）"
					Response.End()
				End Select
				PassStr = PassStr & AnsiStr
			Next
			   PassWordFromBT = PassWordFromBT & Cstr(Chr(Cint(PassStr)))
		Next
	Else
		Response.Write "解密内容为空，解密失败！（操作中断！）"
		Response.End()
	End If
End Function

'加密JX---畸型，HOHO……
Function PassWordToJX(pswstr)
	Dim Ci,Cj,PassStr,PwsAsc,PswAscToStr,PswAscStr,AnsiStr
	PassWordToJX = ""
	pswstr = Replace(pswstr," ","")
	If trim(pswstr) <> "" Then
		For Ci = 1 To Len(PswStr)
			PassStr = ""
			PswAscStr = Cstr(Asc(Lcase(Mid(PswStr,Ci,1))))
			For Cj = 1 To Len(PswAscStr)
				PswAscToStr=Mid(PswAscStr,Cj,1)
				Select Case PswAscToStr
				   Case "1":AnsiStr = "√"
				   Case "2":AnsiStr = "←"
				   Case "3":AnsiStr = "○"
				   Case "4":AnsiStr = "♂"
				   Case "5":AnsiStr = "■"
				   Case "6":AnsiStr = "※"
				   Case "7":AnsiStr = "×"
				   Case "8":AnsiStr = "●"
				   Case "9":AnsiStr = "♀"
				   Case "0":AnsiStr = "◎"
				End Select
				PassStr = PassStr & AnsiStr
			 Next
			 If PassWordToJX = "" Then
				PassWordToJX = PassStr
			 Else
				PassWordToJX = PassWordToJX & "≌" & PassStr
			 End If
		Next
	Else
		Response.Write "加密内容为空，加密失败！（操作中断！）"
		Response.End()
	End If
End Function

'解密JX---畸型，HOHO……
Function PassWordFromJX(pswstr)
	Dim Ci,Cj,PassStr,PasAsc,PswAscToStr,PswAscStr,AnsiStr
	PassWordFromJX = ""
	If trim(pswstr) <> "" Then
		PasAsc = split(pswstr,"≌")
		For Ci = 0 To Ubound(PasAsc)
			PassStr = ""
			PswAscStr = PasAsc(Ci)
			For Cj = 1 To Len(PswAscStr)
				PswAscToStr = Mid(PswAscStr,Cj,1)
				Select Case PswAscToStr
				  Case "√":AnsiStr = "1"
				  Case "←":AnsiStr = "2"
				  Case "○":AnsiStr = "3"
				  Case "♂":AnsiStr = "4"
				  Case "■":AnsiStr = "5"
				  Case "※":AnsiStr = "6"
				  Case "×":AnsiStr = "7"
				  Case "●":AnsiStr = "8"
				  Case "♀":AnsiStr = "9"
				  Case "◎":AnsiStr = "0"
				  Case Else:
					Response.Write "解密内容含非法字符，解密失败！（操作中断！）"
					Response.End()
				End Select
				PassStr = PassStr & AnsiStr
			Next
			   PassWordFromJX = PassWordFromJX & Cstr(Chr(Cint(PassStr)))
		Next
	Else
		Response.Write "解密内容为空，解密失败！（操作中断！）"
		Response.End()
	End If
End Function

'取单双引号
Function ReplaceQM(Str)
	If Str = "" Then
		ReplaceQM = ""
		Exit Function
	End If
	Str = replace(Str,"'","")
	Str = replace(Str,chr(34),"")
	Str = replace(Str,"/**/","")
	Str = replace(Str,"char(0x27)","")
	Str = replace(Str,";","")
	'Str = replace(Str,",","")
	'Str = replace(Str,".","")
	ReplaceQM = Str
End Function



function returnTextarea(str)
	dim temp
	if isnull(str) or str = "" then
		returnTextarea = "&nbsp;"
	else
		temp = str
		temp = replace(temp,vbcrlf,"<br>")
		temp = replace(temp,"<br />","<br>")
		temp = replace(temp," ","&nbsp;")
		returnTextarea = temp
	end if
end function

function showMaxNum(num,maxNum)
	dim i,N
	N = cstr(num)
	for i = len(N) to MaxNum
		N = "0" & N
	next
	showMaxNum = N
end function

'//////////////////////以下函数非通用函数////////////////////////////////////////

Dim UserIp
'客户端IP
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Then
	UserIp = Request.ServerVariables("REMOTE_ADDR")
Else
	UserIp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
End If

function showDate(isStat,weekNow)
	if isStat = true then
		showDate = DateAdd("d",(weekNow - 1) * 7,year(now()) & "-1-1")
	else
		showDate = DateAdd("d",(weekNow) * 7 - 1,year(now()) & "-1-1")
	end if
	if len(day(showDate)) = 1 then
		showDate = year(showDate) & "-" & month(showDate) & "-0" & day(showDate)
	end if
end function


Function weekNo()
	dim dayCountNow,weekNow
	dayCountNow = datediff("d",year(now()) & "-1-1",date()) + 1
	weekNow = dayCountNow / 7
	if weekNow > fix(weekNow) then
		weekNow = fix(weekNow) + 1
	end if
	
	weekNo = weekNow
End Function 

sub WhereTable(pic,title)
%>
<table class='header' border='0' cellspacing='0' cellpadding='0' width='100%'>
<tr>
<td><img border='0' src='images/<%=pic%>' align='absmiddle' width='32' height='32'> <font style='font-size:16pt;'><%=title%></font></td>
<td align='right'>&nbsp;</td>
<td width='100' align='right'> <a onclick='location.reload();'><img style='cursor:hand;' alt='刷新本页面' border='0' src='images/reload.png'></a>  </td>
</tr>
</table>
<%
end sub

dim getallxiashustr
sub getallxiashu(a1)
set rsu=conn.execute("select id from dept_login where id_highdept="&a1&"")
if not rsu.eof then
	do while not rsu.eof
		getallxiashustr=getallxiashustr&" or id_zrr in (select id from login where id_dept="&rsu(0)&")"
		call getallxiashu(rsu(0))
	rsu.movenext
	loop 
end if
rsu.close
end sub

Sub CloseConn()
	On Error Resume Next
	Conn.Close
	Set Conn = Nothing
End Sub
%>