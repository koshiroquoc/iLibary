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
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<form method="POST" name="frmSearch" action="default.asp?name=bookresult">
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10">
						<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
						<td width="228"><b><a href="default.asp?name=catebook">Quay lại</a></b></td>
						<td width="67"><b>Tìm nhanh</b></td>
						<td>
						<p align="center" style="margin-right: 3px">
						<input type="text" name="txtSearchKey" size="14" class="textbox"></td>
						<td width="39">
						<button name="B1" type="submit" class="input_button">&nbsp;Tìm
						</button></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
		<input type="hidden" name="category" value="searchbook">
		</form>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<%					
					strSQL = "SELECT * FROM CATEGORY_BOOK ORDER BY NAME"
					Set rsListBook = Server.CreateObject("ADODB.Recordset")
					rsListBook.Open strSQL,Conn,3,1					
				%>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%
						If rsListBook.Eof Then
					%>
					<tr>
						<td colspan="4">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						Không có file trong mục này</td>
					</tr>
					<%
						Else
					%>
					<tr>
					<%	
						iCount = 1
						Do While Not rsListBook.Eof 
						If iCount Mod 2 <> 0 Then
					%>						
						<td width="25%">
						<p align="right" style="margin-right: 4px; margin-top:4px; margin-bottom:4px; line-height:200%">
						<img border="0" src="../images/pic/book1.ico" width="25" height="25"></td>
						<td width="27%"><a href="default.asp?name=listcatebook&id=<%=rsListBook("CATEGORY_ID")%>"><%=rsListBook("NAME")%></a></td>
					<%
						Else
					%>	
						<td width="16%">
						<p style="margin-right: 4px; margin-top:4px; margin-bottom:4px; line-height:200%" align="right">
						<img border="0" src="../images/pic/book1.ico" width="25" height="25"></td>
						<td width="31%"><a href="default.asp?name=listcatebook&id=<%=rsListBook("CATEGORY_ID")%>"><%=rsListBook("NAME")%></a></td>
					</tr>
					<%
						End If
						iCount = iCount + 1
						rsListBook.MoveNext
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
		<td>&nbsp;</td>
	</tr>
</table>

</div>
</body>

</html>