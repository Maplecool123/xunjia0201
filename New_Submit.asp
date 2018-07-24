<!--#include file="Index_top.asp" -->
<%
if session("iflogin")-99=0 then
if (request.form("action")="del2") then
	if request.form("selectdel")<>"" then
		For i = 1 To Request.Form("selectdel").Count
			SelectTable=Request.Form("selectdel")(i)
			select case SelectTable
			'初始化公告表
			case "gonggao" conn.execute "delete from gonggao"
			'初始化总栏表
			case "yijian" conn.execute "delete from yijian"
			'初始化大栏表
			case "rizi" conn.execute "delete from rizi"
			'初始化小栏表
			case "companygrade" conn.execute "delete from companygrade"
			'初始化专题表
			case "companytype" conn.execute "delete from companytype"
			'初始化文章表
			case "company"
			set rs2=conn.execute("select * from company")
			if not rs2.eof then
				do while not rs2.EOF
				companyfile=rs2("companyfile")
				companycontract=rs2("companycontract")
				
				if companyfile<>"" then
					companyfiles=split(companyfile,",")
					for ii=0 to ubound(companyfiles)
						Set fso=Server.CreateObject("Scripting.FileSystemObject")
						If fso.FileExists(Server.Mappath("Uploadfile/companyfile/"&companyfiles(ii)))=true Then
							fso.DeleteFile Server.Mappath("Uploadfile/companyfile/"&companyfiles(ii))
						End If
						Set fso=Nothing
					next
				end if
				
				if companycontract<>"" then
					companycontracts=split(companycontract,",")
					for ii=0 to ubound(companycontracts)
						Set fso=Server.CreateObject("Scripting.FileSystemObject")
						If fso.FileExists(Server.Mappath("Uploadfile/companycontract/"&companycontracts(ii)))=true Then
							fso.DeleteFile Server.Mappath("Uploadfile/companycontract/"&companycontracts(ii))
						End If
						Set fso=Nothing
					next
				end if
				rs2.movenext
				loop
			end if
			rs2.close
			set rs2=nothing
			
			conn.execute "delete from company"
			'初始化临时表
			case "tender" conn.execute "delete from tender"
			'初始化上传临时表
			case "competitive" conn.execute "delete from competitive"
			end select
		Next
				
		call HintAndTurn("相关数据删除成功！","New.asp")
		response.end
	end if
end if

if (request.form("action")="del1") then
	conn.execute "delete from gonggao"
	conn.execute "delete from yijian"
	conn.execute "delete from rizi"
	conn.execute "delete from companygrade"
	conn.execute "delete from companytype"
			
	set rs2=conn.execute("select * from company")
	if not rs2.eof then
		do while not rs2.EOF
		companyfile=rs2("companyfile")
		companycontract=rs2("companycontract")
				
		if companyfile<>"" then
			companyfiles=split(companyfile,",")
			for ii=0 to ubound(companyfiles)
				Set fso=Server.CreateObject("Scripting.FileSystemObject")
				If fso.FileExists(Server.Mappath("uploadfile/companyfile/"&companyfiles(ii)))=true Then
					fso.DeleteFile Server.Mappath("uploadfile/companyfile/"&companyfiles(ii))
				End If
				Set fso=Nothing
			next
		end if
				
		if companycontract<>"" then
			companycontracts=split(companycontract,",")
			for ii=0 to ubound(companycontracts)
				Set fso=Server.CreateObject("Scripting.FileSystemObject")
				If fso.FileExists(Server.Mappath("uploadfile/companycontract/"&companycontracts(ii)))=true Then
					fso.DeleteFile Server.Mappath("uploadfile/companycontract/"&companycontracts(ii))
				End If
				Set fso=Nothing
			next
		end if
		rs2.movenext
		loop
	end if
	rs2.close
	set rs2=nothing
			
	conn.execute "delete from company"
	conn.execute "delete from tender"
	conn.execute "delete from competitive"
				
	call HintAndTurn("所有数据删除成功！","New.asp")
	response.end
end if

else
response.redirect "erro.asp"
response.end
end if%>
<!--#include file="Index_bottom.asp" -->