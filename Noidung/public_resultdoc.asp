<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
	<%
		txtSearchKey = Request.Form("txtSearchKey")
				
		strSQL = "SELECT * FROM DOCUMENT WHERE TITLE LIKE '%" & txtSearchKey & "%'"
		Set rsDocument = Server.CreateObject("ADODB.Recordset")
		rsDocument.Open strSQL,Conn,3,1					
	%>				
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" id="table2" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
						<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table3">
							<tr>
								<td width="10">
								<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
								<td width="304"><b>
								<a href="default.asp?name=listdoc">Quay lại</a></b></td>
								<td>
								<p align="right" style="margin-right: 3px">
								&nbsp;</td>
							</tr>
						</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%
						If rsDocument.Eof Then
					%>
					<tr>
						<td>
				<p align="center" style="margin-top: 10px; margin-bottom:10px">
				Không tìm thấy tài liệu nào trong mục này<br>
				Bấm <a href="JavaScript:history.go(-1)">vào đây</a> để quay lại.</td>
					</tr>
					<%
						Else
					%>
					<tr>
						<td>
						<p align="center" style="margin-top: 3px; margin-bottom: 3px">
						<font color="#0066FF"><b>Có tổng số
						 <%
							If rsDocument.RecordCount > 6 Then
								Response.Write 6
							Else
								Response.Write rsDocument.RecordCount
							End If
					 	%> kết quả tìm được</b></font></td>
					</tr>
					<%
						End If
					%>	
					<tr>
						<td>
						<img border="0" src="../images/spacer.gif" width="1" height="3"></td>
					</tr>
					<%
						If Not rsDocument.Eof Then
						iDocument = 1
						Do While Not rsDocument.Eof And iDocument <=3
					%>
					<tr>
						<td>
						<p style="margin-left:5px; margin-right:6px; margin-top:1px; margin-bottom:1px">
						<a href="default.asp?name=docdetail&id=<%=rsDocument("DOCUMENT_ID")%>"><b><%=rsDocument("TITLE")%></b></a></td>
					</tr>
					<tr>
						<td>
						<p style="margin-left:5px; margin-right:6px; margin-top:1px; margin-bottom:1px" align="justify"><%=rsDocument("SUMMARY")%></td>
					</tr>
					<tr>
						<td>
						<p align="center">&nbsp;</td>
					</tr>
					<%
						iDocument = iDocument + 1
						rsDocument.MoveNext
						Loop
						End If
						Conn.Close
						Set Conn = Nothing
					%>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<p align="center">
		<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
	</tr>
</table>

</div>
</body>

</html>