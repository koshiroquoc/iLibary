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
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10">
						<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
						<td width="237"><b><a href="default.asp?name=listsoft">Quay lại</a></b></td>
						<form method="POST" name="frmList">
						<td width="75">
						<p align="center">
						<span style="font-size: 8pt; font-weight: 700">Tìm nhanh</span></td>
						<td>
						<p align="center" style="margin-right: 3px">
						<input type="text" name="txtSoftName" size="11" class="textbox"></td>
						<td width="40">
						<button name="B1" type="submit" class="input_button">&nbsp;Tìm
						</button></td>
						</form>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<%
					id	= Request.QueryString("id")
					strSQL = "SELECT * FROM SOFTWARE WHERE CATEGORY_ID="& id
					strSQL = strSQL & " ORDER BY NAME"
					Set rsListSoft = Server.CreateObject("ADODB.Recordset")
					rsListSoft.Open strSQL,Conn,3,1					
				%>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%
						If rsListSoft.Eof Then
					%>
					<tr>
						<td colspan="4">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						Không tìm thấy dữ liệu này, xin nhập lại!</td>
					</tr>
					<%
						Else
					%>
					<tr>
					<%	
						iCount = 1
						Do While Not rsListSoft.Eof 
						If iCount Mod 2 <> 0 Then
					%>						
						<td width="58">
						<p align="center" style="margin-right: 6px; margin-top:6px; margin-bottom:6px">
						<img border="0" src="../images/Pic/saveitem.gif" width="16" height="16" align="right"></td>
						<td width="173"><b><%=rsListSoft("NAME")%></b><br>
						<a href="<%=rsListSoft("FILE_PATH")%>">Download</a></td>
					<%
						Else
					%>	
						<td width="62">
						<p style="margin-right: 6px; margin-top:6px; margin-bottom:6px" align="right">
						<img border="0" src="../images/Pic/saveitem.gif" width="16" height="16" align="right"></td>
						<td><b><%=rsListSoft("NAME")%></b><br>
						<a href="<%=rsListSoft("FILE_PATH")%>">Download</a></td>
					</tr>
					<%
						End If
						iCount = iCount + 1
						rsListSoft.MoveNext
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