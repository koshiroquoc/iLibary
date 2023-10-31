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
		strSQL = "SELECT * FROM NOTICE ORDER BY DATE_INFORM DESC"
		Set rsNotice = Server.CreateObject("ADODB.Recordset")
		rsNotice.Open strSQL,Conn,3,1
	%>
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10">
						<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
						<td width="207"><b><a href="default.asp">Quay lại</a></b></td>
						<td width="69"><b><font size="2">Tìm nhanh</font></b></td>
						<form method="POST" name="frmList" action="default.asp?name=resultnotice">
						<td width="204">
						<p align="right">
						<font size="2">
						<input name="txtSearchKey" size="35" class="textbox" style="float: left"></font></td>
						<td>
						<p align="center" style="margin-right: 3px; margin-top:1px">
						<button name="B1" class="input_button" type="submit">
						<font size="1">&nbsp;Tìm&nbsp;</font>
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
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<%
					iCount = 1
					Do While Not rsNotice.Eof And iCount <=5
				%>
					<tr>
						<td>
						<p style="margin-left: 4px; margin-right: 2px; margin-top: 4px; margin-bottom: 2px">
						<img src="../images/blackArrow_right.gif"><b><a href="default.asp?name=notidetail&id=<%=rsNotice("ID")%>"><%=rsNotice("TITLE")%>&nbsp; -&nbsp; Ngày <%=NgayVN(rsNotice("DATE_INFORM"))%></a></b>
						<% If (Now() - rsNotice("DATE_INFORM") < 3) Then %>
						<img border="0" src="../images/new.gif" width="33" height="16" align="middle">
						<% End If %>
						</td>
					</tr>
					<tr>
						<td>
						<p style="margin-left:4px; margin-right:4px; margin-top:2px; margin-bottom:2px" align="justify"><%=rsNotice("SUMMARY")%></td>
					</tr>
					<%
						If iCount <> rsNotice.RecordCount Then
					%>
					<tr>
						<td>
						<p align="center" style="margin-top: 6px">
						<img border="0" src="../images/line.gif" width="300" height="3"></td>
					</tr>
					<%
						Else
					%>
					<tr>
						<td>
						<img border="0" src="../images/spacer.gif" width="1" height="6"></td>
					</tr>
					<%
						End If
						iCount = iCount + 1
						rsNotice.MoveNext
						Loop
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