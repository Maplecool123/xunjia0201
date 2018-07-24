<!--#include file="upload_5xsoft.inc"-->
<!--#include file="conn.asp"-->
<!--#include file="const1.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/uploadcontrol.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<%
gourl=""
set upload=new upload_5xsoft
id=upload.form("id")
gotourl=upload.form("gotourl")
companyrealname=upload.form("companyrealname")

actiondetail=upload.form("actiondetail")
allpaixu=""
havecompanyfile=""
i=0
	
if actiondetail="saveedit" then
	gourl="company_class="&upload.form("company_class")&"&company_type="&upload.form("company_type")&"&checkstatue="&upload.form("checkstatue")&"&checkcondition="&upload.form("checkcondition")&"&checkstyle="&upload.form("checkstyle")&"&keyword1="&upload.form("keyword1")&"&page="&upload.form("page")
	picno=upload.form("picno")
	if picno>=1 then
		for nnn=1 to picno
		if upload.form("havecompanyfile"&nnn)<>"" then
			havecompanyfile=havecompanyfile&upload.form("havecompanyfile"&nnn)&","
			allpaixu=allpaixu&upload.form("allpaixu"&nnn)&","
			i=i+1
		end if
		next
	end if
	
	for nnn=1 to maxcfile
		if upload.form("havecompanyfile"&nnn)="" then
			allpaixu=allpaixu&upload.form("allpaixu"&nnn)&","
		end if
	next
	
	if allpaixu<>"" then allpaixu=left(allpaixu,len(allpaixu)-1)
end if

if upload.form("action")="uploadfile" then
	filepath="uploadfile\companyfile\"
	if instr(filepath,"*")>0 or instr(filepath,"?")>0 or instr(filepath,"'")>0 or instr(filepath,chr(34))>0 then
		msg="对不起，没有填写路径或路径中含有非法字符<br><br>"
		HX_GoBack msg,Url
	end if
	if right(filepath,1)<>"\" then
		filepath=filepath&"\"
		filepath=PathReplace(filepath)
	end if
	'response.write filepath
	'response.end
	basepath=server.MapPath(trim(filepath))
	set fso=server.createobject("scripting.filesystemobject")
	if not fso.folderexists(basepath) then
		set fso=nothing
		set upload=nothing
		call HintAndBack("对不起，目录不存在，请先创建该目录！",1)
	end if
	
	for each formName in upload.objFile
		set file=upload.objFile(formName)
		if file.filename<>"" and isnull(file.filename)=false then
			if file.FileSize>uploadlimit*1024*10240 then
				call HintAndBack("对不起，文件大小不可以超过"&uploadlimit&"M！",1)
				exit for
			end if
		end if
		set file=nothing
	next
	
	uploadflag="no"
	for each formName in upload.objFile
		set file=upload.objFile(formName)
		if file.filename<>"" and isnull(file.filename)=false then
		uploadflag="yes"
		exit for
		end if
		set file=nothing
	next
	'response.end
	if uploadflag="no" and actiondetail<>"saveedit" then
		call HintAndBack("请选择至少选择一个上传的文件！",1)
	end if
	
	allpicstr=""
	for each formName in upload.objFile
		set file=upload.objFile(formName)
		if file.filename<>"" and isnull(file.filename)=false then
			filename=companyrealname&"_"&file.FileName
			filesize=file.FileSize
			fileExt=trim(lcase(filename))
			fileExt=split(fileExt,".")(ubound(split(fileExt,".")))
			uploadsuc=false
			for ii=0 to ubound(Forumupload)
				if fileEXT=trim(Forumupload(ii)) then
					uploadsuc=true
					exit for
					'call HintAndBack("文件格式不正确,重新上传！",1)
				'else
					'uploadsuc=false
				end if
			next
			if uploadsuc=false then
				call HintAndBack("文件格式不正确,重新上传！",1)
			end if
			FileSeverPath = basepath&"\"&FileReplace(filename)
				
			'response.write FileSeverPath
			'response.end
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			if fso.FileExists(FileSeverPath) then
				'首先去掉后缀名
				nohouzhui=""
				houzhui=""
				if instr(FileReplace(filename),".")>0 then   '如果有后缀名
					nohouzhuis=split(FileReplace(filename),".")
					for ii=0 to ubound(nohouzhuis)-1
						nohouzhui=nohouzhui&nohouzhuis(ii)&"."
					next
					nohouzhui=left(nohouzhui,len(nohouzhui)-1)
					houzhui="."&nohouzhuis(ubound(nohouzhuis))
				else
					nohouzhui=filename
				end if
					
				for ii=1 to 99
					filename=nohouzhui&"副本"&ii&houzhui
					FileSeverPath = basepath&"\"&FileReplace(filename)
					if not fso.FileExists(FileSeverPath) then
					exit for
					end if
				next
			end if	
			set fso=nothing
			if filesize>0 then
				file.SaveAs FileSeverPath
				'call WS_S.DirFileUpdateEx (prefolder,FileReplace(filename),filesize,now(),nowusername,uploadman,filetype)
				'call writetxt (session("admin_realname")&"（"&session("admin_name")&"）","上传了文件",FileSeverPath)
				allpicstr=allpicstr&FileReplace(filename)&","
				i=i+1
			end if
			if ScanFile(FileSeverPath) then
				CheckFile FileSeverPath,True 
				call HintAndBack("对不起,您上传的文件中含有木马程序,所以不能上传！",1)
			end if
		end if
		set file=nothing
	next
	
	allpicstr=havecompanyfile&allpicstr
	
	if allpicstr<>"" and right(allpicstr,1)="," then
		allpicstr=left(allpicstr,len(allpicstr)-1)
		allpicstrs=split(allpicstr,",")
		allpaixus=split(allpaixu,",")
		
		redim arr(ubound(allpicstrs),2)
		
		for i=0 to ubound(allpicstrs)
			arr(i,1)=allpicstrs(i)
			arr(i,2)=allpaixus(i)
		next
		
		tmp1=""
		tmp2=""
		For i = 0 To ubound(allpicstrs)  '从小到大排列
			For j = i + 1 To ubound(allpicstrs)
				If arr(i,2)-arr(j,2)>0 Then
					tmp1 = arr(i,2)
					arr(i,2) = arr(j,2)
					arr(j,2) = tmp1
					tmp2 = arr(i,1)
					arr(i,1) = arr(j,1)
					arr(j,1) = tmp2
				End If
			Next
		Next
		
		allpicstr=""
		allpaixu=""
		For i = 0 To ubound(allpicstrs)
			allpicstr=allpicstr&arr(i,1)&","
			allpaixu=allpaixu&arr(i,2)&","
		Next
		
		'response.write "<br><br>--"&allpicstr
		'response.write "<br><br>--"&allpaixu
		'response.end
		allpicstr=left(allpicstr,len(allpicstr)-1)
		
		'response.write allpicstr
		'response.end
		
		set rs=server.createobject("ADODB.RecordSet")
		sql="select * from company where id="&id
		rs.open sql,conn,1,3
		rs("companyfile")=allpicstr
		if rs("regstatue")<=2 then
			rs("regstatue")=2
		end if
		rs.update
		rs.close
	end if
	
	'response.write allpaixu
	'response.end
	call HintAndTurn("共"&i&"个文件上传成功，下一步上传企业合同材料！",gotourl&".asp?id="&id&"&companyrealname="&companyrealname&"&"&gourl)
	set upload=nothing
else
	set upload=nothing
	call HintAndBack("非法操作！",1)
end if%>
