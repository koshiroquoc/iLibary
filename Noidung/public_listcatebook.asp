<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
		id	= Request.QueryString("id")
		If id <> "" Then
			strSQL = "SELECT * FROM BOOK WHERE LEFT(BOOK_ID,3)='" & id & "'"	
		Else
			strSQL = LoadSQL()
		End If
			
		Set rsListBook = Server.CreateObject("ADODB.Recordset")
		rsListBook.Open strSQL,Conn,3,1
		InsertSQL(strSQL)
		If rsListBook.Eof Then		
	%>
	<tr>
		<td>
		<p align="center"><b>Không tồn tại sách trong mục bạn chọn! <br>
		Kích <a href="JavaScript:history.back();">vào đây</a> để quay lại!</b></td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="20"></td>
	</tr>
	<%
		Else		  	
		  	If Request.QueryString("page") = "" Then
				intCurrentPage = 1
			Else
				intCurrentPage = CInt(Request.QueryString("page"))
			End If
			rsListBook.PageSize = 4
			If rsListBook.PageCount > 0 then
				rsListBook.AbsolutePage = intCurrentPage
			Else
				intCurrentPage = 0
			End If
			numPage = rsListBook.PageCount
			numResult = rsListBook.RecordCount
	%>
	<tr>
		<td>
		<p align="center">Có tất cả <%=numResult%> cuốn sách trong danh mục này.</td>
	</tr>
	<tr>
		<td>
		<p align="center">
		<img border="0" src="../images/spacer.gif" width="1" height="10"></td>
	</tr>
		<%
			iCount = 1
			Do While Not rsListBook.Eof and iCount <= rsListBook.PageSize
			strSQL = "SELECT * FROM CATEGORY_BOOK WHERE CATEGORY_ID ='" & Left(rsListBook("BOOK_ID"),3) & "'"
			Set rsCategoryBook = Server.CreateObject("ADODB.Recordset")
			rsCategoryBook.Open strSQL,Conn,3,1
			
			strSQL = "SELECT * FROM PUBLISHER WHERE ID=" & rsListBook("PUBLISHER")
			Set rsPublisher = Server.CreateObject("ADODB.Recordset")
			rsPublisher.Open strSQL,Conn,3,1
			
			strSQL = "SELECT * FROM LANGUAGE WHERE ID=" & rsListBook("LANGUAGE")
			Set rsLanguage = Server.CreateObject("ADODB.Recordset")
			rsLanguage.Open strSQL,Conn,3,1
		%>
	<tr>
		<td>
		<table border="1" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 8px; margin-bottom: 2px"><b>Tên 
						sách</b></td>
						<td width="2%">
						<p style="margin-top: 8px"><b>:</b></td>
						<td width="56%"><font color="#0066FF"><b><%=rsListBook("NAME")%></b></font></td>
						<td width="20%" rowspan="6">
						<p style="margin-top: 8px">
						<img border="0" width="81" height="91" align="right" src="<%=rsListBook("IMAGE")%>"></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Mã sách</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsListBook("BOOK_ID")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Lĩnh vực</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsCategoryBook("NAME")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<b>Tên tác giả</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsListBook("AUTHOR")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Nhà xuất bản</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsPublisher("NAME")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Ngôn ngữ</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsLanguage("NAME")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 8px"><b>Tóm tắt</b></td>
						<td width="2%">
						<p style="margin-bottom: 8px"><b>:</b></td>
						<td colspan="2">
						<p align="justify" style="margin-bottom: 8px"><%=rsListBook("SUMMARY")%></td>
						<td width="2%">&nbsp;</td>
					</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
		<%
			iCount = iCount + 1
			rsListBook.MoveNext
			Loop
			Conn.Close
			Set Conn = Nothing
		%>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
		If numPage >1 Then
	%>
	<tr>
		<td>
		<table border="1" width="100%" id="table3" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
			<tr>
				<td bgcolor="#EEEEEE">
				<table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
					<tr>
						<td width="18"><a href="JavaScript:history.back();">
						<img border="0" src="../images/left.gif" width="22" height="20" alt="Quay lại"></a></td>
						<td>&nbsp;</td>
						<td width="20"><a href="#">
						<img border="0" src="../images/top.gif" width="20" height="22" alt="Lên trên"></a></td>
						<td width="386">
							<p align="right">Trang: 
								<%
									for i=1 to numPage
								%>	
								<b><a href="default.asp?name=listcatebook&page=<%=i%>"><%=i%></a> |</b>
								<%	
									Next
								%>						
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<%
		End If
	%>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
		End If
	%>
</table>

</div>
</body>

</html>