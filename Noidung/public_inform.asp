<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Trang chủ</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<table border="0" width="155" cellspacing="0" cellpadding="0">
	<tr>
		<td>
	<%
		Set rsNotice = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM NOTICE ORDER BY DATE_INFORM DESC"
		rsNotice.Open strSQL,Conn,3,1
		If Not rsNotice.Eof Then
	%>
	<table cellpadding="0" cellspacing="0" style="border-collapse: collapse"  width="155" id="table5" bgcolor="#FFFFFF">
		<tr>
			<td height="20" width="20" style="border-left: 1px solid #999999; border-top: 1px solid #999999" bgcolor="#3399FF">
			<font color="#FFFFFF">
			<b>&nbsp;..::</b></font></td>
			<td height="20" width="155" style="border-right: 1px solid #999999; border-top: 1px solid #999999" class="txt_titlemenu" bgcolor="#3399FF">
			<font color="#FFFFFF"><b>THÔNG BÁO</b></font></td>
		</tr>
		<tr>
			<td width="100%" colspan="2">
				<table border="0" cellpadding="5" cellspacing="0"  width="100%" id="table6">
					<tr>
						<td style="border:0px solid #999999; " valign="top" height="130">
							<marquee onmouseover=this.stop() onmouseout=this.start() scrollAmount=2 scrollDelay=120 direction=up style="text-align: justify; " height=138>
							<%
								iCount = 1
								Do While Not rsNotice.Eof And iCount <=5
							%>
							<p style="margin-top: 3px; margin-bottom: 3px">
							<a href="default.asp?name=notidetail&id=<%=rsNotice("ID")%>"><%=Trim(rsNotice("SUMMARY"))%></a><br>
							<%
								iCount = iCount + 1
								rsNotice.MoveNext
								Loop
							%>
							</marquee>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="2"></td>
	</tr>
<%
	End if
%>
	<tr>
		<td>
		<table cellpadding="0" cellspacing="0" style="border-collapse: collapse"  width="155" id="table7" bgcolor="#FFFFFF">
			<tr>
				<td height="21" width="20" style="border-left: 1px solid #999999; border-top: 1px solid #999999" bgcolor="#3399FF">
				<font color="#FFFFFF">
			<b>&nbsp;..::</b></font></td>
				<td height="21" width="155" style="border-right: 1px solid #999999; border-top: 1px solid #999999" class="txt_titlemenu" bgcolor="#3399FF">
				<font color="#FFFFFF"><b>SÁCH MỚI</b></font></td>
			</tr>
			<tr>
				<td width="100%" colspan="2">
					<table border="0" cellpadding="5" cellspacing="0"  width="100%" id="table8">
					<tr>
						<td style="border-style:solid; border-width:0px; " valign="top" height="160">
						<%
							Set rsBook = Server.CreateObject("ADODB.Recordset")
							strSQL = "SELECT * FROM BOOK ORDER BY DATE_INFORM DESC"
							rsBook.Open strSQL,Conn,3,1
							If Not rsBook.Eof Then
						%>		
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<%
							iBook = 1
							Do While Not rsBook.Eof and iBook <10
							If rsBook("IMAGE")<>"" Then
						%>
						<p style="margin-top: 3px; margin-bottom: 3px">
							<tr>
								<td ><span style="font-size: 10pt">
								<p style="text-align: left">
								<img border="0" src="../images/new.gif"><a href="default.asp?name=homebook&id=<%=rsBook("ID")%>"><%=rsBook("NAME")%></a></span></td>
							</tr>
							<tr>
								<td valign="top" class="book_title" height="4"></td>
							</tr>
						<%
							Else
						%>	
							<tr>
								<td ><span style="font-size: 10pt"><%=rsBook("NAME")%></span> </td>
							</tr>
						<%
							End If
							iBook = iBook + 1
							rsBook.MoveNext
							Loop
						%>	
						</table>
						<%
							End If
							Conn.Close
							Set Conn = Nothing
						%>
						</td>
					</tr>
					</table>					
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="2"></td>
	</tr>
</table>
</body>

</html>