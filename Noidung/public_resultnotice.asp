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
		txtSearchKey = Request.Form("txtSearchKey")
		strSQL = "SELECT * FROM NOTICE WHERE TITLE LIKE '%" & txtSearchKey & "%'"
		Set rsNotice = Server.CreateObject("ADODB.Recordset")
		rsNotice.Open strSQL,Conn,3,1
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
						<td width="205"><b><a href="default.asp?name=notice">Quay lại</a></b></td>
						<td width="71"><b>Tìm nhanh</b></td>
						<form method="POST" name="frmList" action="default.asp?name=resultnotice">
						<td width="104">
						<p align="center">
						<input name="txtSearchKey" size="16" class="textbox" style="float: left"></td>
						<td>
						<p align="center" style="margin-right: 3px; margin-top:1px">
						<button name="B1" class="input_button" type="submit">
						&nbsp;Tìm&nbsp;
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
	<%
		If rsNotice.Eof Then
	%>
	<tr>
		<td>
		<table border="1" width="100%" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC" cellspacing="0" cellpadding="0">
			<tr>
				<td>
				<p align="center" style="margin-top: 10px; margin-bottom:10px">
				Không tìm thấy tài liệu giống khóa tìm kiếm<br>
				Bấm <a href="JavaScript:history.go(-1)">vào đây</a> để quay lại.</td>
			</tr>
		</table>
		</td>
	</tr>
	<%
		Else
	%>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%
						iCount = 1
						Do While Not rsNotice.Eof And iCount <=7
					%>
					<tr>
						<td>
						<p style="margin-left: 4px; margin-right: 2px; margin-top: 4px; margin-bottom: 2px">
						<b>
						<a href="default.asp?name=notidetail&id=<%=rsNotice("ID")%>"><%=rsNotice("TITLE")%>&nbsp; -&nbsp; Ngày <%=NgayVN(rsNotice("DATE_INFORM"))%></b></a>
						<% If (Now() - rsNotice("DATE_INFORM") < 3) Then %>
						<img border="0" src="../images/new.gif" width="33" height="16">
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
					%>
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
		Conn.Close
		Set Conn = Nothing
	%>
</table>

</div>
</body>

</html>