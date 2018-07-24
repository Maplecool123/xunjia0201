<!--#include file="Conn.asp" -->
<!--#include file="Inc/Function.asp" -->
<%sum=0

set rs1=conn.execute("select djh from competitive group by djh")
if not rs1.eof then
	do while not rs1.eof
	
	set rs4=server.createobject("ADODB.RecordSet")
	sql4="select needdate,statue,companyno,enddate,enddatehour from tender where djh='"&rs1("djh")&"' and ifzu=1"
	rs4.open sql4,conn,1,3
	if not rs4.eof then
		needdate=rs4("needdate")
		statue=checkifnum(rs4("statue"))
		companyno=checkifnum(rs4("companyno"))
		enddate=rs4("enddate")
		enddatehour=rs4("enddatehour")
	end if
	rs4.close
	
	set rs2=conn.execute("select companyid from competitive where djh='"&rs1("djh")&"' group by companyid")
	if not rs2.eof then
		do while not rs2.eof
			set rs=server.createobject("ADODB.RecordSet")
			sql="select * from competitive where djh='"&rs1("djh")&"' and companyid="&rs2("companyid")&" and ifzu=1"
			rs.open sql,conn,1,3
			if rs.eof then
			
				totalprice=0
				spotprice=0
				set rs3=conn.execute("select sum(totalprice),sum(spotprice) from competitive where djh='"&rs1("djh")&"' and companyid="&rs2("companyid")&"")
				if not rs3.eof then
					totalprice=checkifnum(rs3(0))
					spotprice=checkifnum(rs3(1))
				end if
				rs3.close
				
				addtime=shownamestrnull("addtime","competitive","djh",rs1("djh")," and companyid="&rs2("companyid")&"")
				
				sum=sum+1
				rs.addnew
				rs("djh")=rs1("djh")
				rs("tenderid")=0
				rs("companyid")=rs2("companyid")
				rs("singleprice")=0
				rs("totalprice")=totalprice
				rs("spotprice")=spotprice
				rs("deliverydate")=needdate
				'rs("paystyle")=nowpaystyle
				'rs("beizhu")=""
				rs("addtime")=addtime
				rs("statue")=0
				rs("ifzu")=1
				rs.update
				response.write sum&"、"&rs("djh")&"-"&shownameint("companyrealname","company","id",rs2("companyid"))&"<br>"
				response.flush
				
			end if
			rs.close
		rs2.movenext
		loop
	end if
	rs2.close
	
	if statue-3=0 then
		companycount=0
		companycount=checkifnum(gettotalitem("competitive","djh",rs1("djh")," and ifzu=1"))
		
		if companycount-companyno>=0 then
			set rs4=server.createobject("ADODB.RecordSet")
			sql4="select * from tender where djh='"&rs1("djh")&"' and ifzu=1"
			rs4.open sql4,conn,1,3
			if not rs4.eof then
				rs4("statue")=2
				rs4.update
			end if
			rs4.close
		end if
	end if
	
	rs1.movenext
	loop
end if
rs1.close

response.write "检查完毕，共计修正"&sum&"条数据。"
response.flush
%>