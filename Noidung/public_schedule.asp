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
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
	</tr>
	<%
		strSQL = "SELECT * FROM SCHEDULE ORDER BY DATE_INFORM DESC"
		Set rsSchedule = Server.CreateObject("ADODB.Recordset")
		rsSchedule.Open strSQL,Conn,3,1
		If Not rsSchedule.Eof Then
	%>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10">
						<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
						<td width="304"><b><a href="default.asp">Quay lại</a></b></td>
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
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="6" height="4"></td>
						<td height="4">
						<p align="center"></td>
						<td width="6" height="4"></td>
					</tr>
					<tr>
						<td width="6">&nbsp;</td>
						<td>
						<p align="center"><b><%=uCase(rsSchedule("TITLE"))%></b></td>
						<td width="6">&nbsp;</td>
					</tr>
					<tr>
						<td width="6" height="4"></td>
						<td height="4">
						<p align="center"></td>
						<td width="6" height="4"></td>
					</tr>
					<tr>
						<td width="6">
						<p style="margin-top: 3px">&nbsp;</td>
						<td><%=rsSchedule("CONTENT")%></td>
						<td width="6">&nbsp;</td>
					</tr>
					<tr>
						<td width="6" height="10"></td>
						<td height="10">
						<p align="center"></td>
						<td width="6" height="10"></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
		End If
		Conn.Close
		Set Conn = Nothing
	%>
</table>

</div>
</body>

</html>