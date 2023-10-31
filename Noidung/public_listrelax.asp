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
		<td><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
	<tr>
		<form method="POST" name="frmList">
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
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
		</form>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="4"></td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<%
					strSQL = "SELECT * FROM CATEGORY_RELAX ORDER BY NAME DESC"
					Set rsCategory= Server.CreateObject("ADODB.Recordset")
					rsCategory.Open strSQL,Conn,3,1					
				%>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%
						If rsCategory.Eof Then
					%>
					<tr>
						<td>
						<p align="center" style="margin-top: 10px; margin-bottom: 10px">Không tìm 
						thấy sách mà bạn đang tra cứu. <br>
		Kích <a href="JavaScript:history.back();">vào đây</a> để quay lại trang tìm kiếm!</td>
					</tr>
					<%
						Else
					%>
					<tr>
					<%	
						iCount = 1
						Do While Not rsCategory.Eof
						strSQL = "SELECT * FROM RELAX WHERE CATEGORY_ID=" & rsCategory("ID")
						strSQL = strSQL & " ORDER BY DATE_INFORM"
						Set rsRelax = Server.CreateObject("ADODB.Recordset")
						rsRelax.Open strSQL,Conn,3,1
						If Not rsRelax.Eof Then
					%>						
						<td width="49%" valign="top">
						<table border="0" width="100%" id="table2" bordercolorlight="#CCCCCC" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
							<tr>
								<td bgcolor="#EEEEEE">
								<p align="center"><b><%=uCase(rsCategory("NAME"))%></b></td>
							</tr>
							<tr>
								<td>
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
									<tr>
										<td>
										<p style="margin-top: 4px; margin-bottom: 0px; margin-left:4px; margin-right:4px">
										<font style="font-size: 10pt" face="Arial">
										<b>
										<a href="default.asp?name=relaxdetail&id=<%=rsRelax("ID")%>"><%=rsRelax("TITLE")%></a></b></font></td>
									</tr>
									<tr>
										<td>
										<p style="margin-top: 2px; margin-bottom: 0px; margin-left:4px; margin-right:4px" align="justify">
										<font style="font-size: 10pt" face="Arial">
										<%=rsRelax("SUMMARY")%></font></td>
									</tr>
									<%
										strSQL = "SELECT * FROM RELAX WHERE CATEGORY_ID=" & rsCategory("ID")
										strSQL = strSQL & " AND ID<>" & rsRelax("ID")
										strSQL = strSQL & " ORDER BY DATE_INFORM"
										Set rsLast = Server.CreateObject("ADODB.Recordset")
										rsLast.Open strSQL,Conn,3,1
										If not rsLast.Eof Then
									%>
									<tr>
										<td>
										<p align="center">
										<img border="0" src="../images/line.gif" width="124" height="5"></td>
									</tr>
									<%
										iIndex = 1
										Do While Not rsLast.Eof And iIndex <=3										
									%>									
									<tr>
									
										<td>
										<font style="font-size: 10pt" face="Arial">
										<p style="margin-left: 4px; margin-right: 4px; margin-top: 4px; margin-bottom: 0px">
										<a href="default.asp?name=relaxdetail&id=<%=rsLast("ID")%>"><img border="0" src="../images/blackArrow_right.gif" width="7" height="5"> <%=rsLast("TITLE")%></a></td>
									</tr>
									<%
										iIndex = iIndex + 1
										rsLast.MoveNext
										Loop								
									End If
									%>
									<tr>
										<td width="100%">
										<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						</td>
						</tr>
					<tr>
						<td width="100%">
						<img border="0" src="../images/spacer.gif" width="1" height="1"></td>
					</tr>
					<%
						End If
						iCount = iCount + 1
						rsCategory.MoveNext
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