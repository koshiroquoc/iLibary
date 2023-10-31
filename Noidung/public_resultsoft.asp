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
				<%
					txtSoftName	= Request.Form("txtSoftName")
					strSQL = "SELECT * FROM SOFTWARE WHERE NAME LIKE '%" & txtSoftName & "%'"
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
				<p align="center" style="margin-top: 10px; margin-bottom:10px">
				Không tìm thấy tài liệu nào trong mục này<br>
				Bấm <a href="JavaScript:history.go(-1)">vào đây</a> để quay lại.</td>
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
						<p align="right" style="margin-right: 6px; margin-top:6px; margin-bottom:6px">
						<img border="0" src="../images/icon_pro.gif" width="40" height="38"></td>
						<td width="173"><b><%=rsListSoft("NAME")%></b><br>
						<a href="<%=rsListSoft("FILE_PATH")%>">Download</a></td>
					<%
						Else
					%>	
						<td width="62">
						<p style="margin-right: 6px; margin-top:6px; margin-bottom:6px" align="right">
						<img border="0" src="../images/icon_pro.gif" width="40" height="38"></td>
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