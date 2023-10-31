<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("book")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>

<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	txtBookID = Request.Form("txtBookID")
	If txtBookID = "" Then
		Response.Redirect("return_error.asp?type=1")
	End If
								
	Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM BORROW WHERE BOOK_ID='" & Trim(txtBookID) & "'"
	rsCheckBorrow.Open strSQL, Conn,3,1
	If rsCheckBorrow.Eof Then			
		rsCheckBorrow.Close
		Set rsCheckBorrow = Nothing
		Response.Redirect("admin_error.asp?type=9")
	Else
			txtBookID = rsCheckBorrow("BOOK_ID")
			txtBookID = rsCheckBorrow("CARD_ID")
		End If
		
		Conn.Close
		Set Conn = Nothing			
	Else
%>					
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body topmargin="8" leftmargin="8">

<div align="center">

<table border="1" width="984" id="table1" bordercolordark="#808080" cellspacing="0" cellpadding="0" bordercolorlight="#D5F1FF">
	<tr>
		<td>
		<div align="center">
			<table border="0" width="984" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td colspan="2"><!--#INCLUDE FILE="admin_header.asp" --></td>
				</tr>
				<tr>
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>&nbsp; 
							<font color="#FF0000" size="2">CẬP NHẬT TRẢ SÁCH</font></b></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="63">&nbsp;</td>
								<td width="379">
								<table border="1" width="100%" id="table12" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="116%" id="table13" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="2">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
											</tr>
											<tr>
												<td colspan="2">
												<p align="center"><b>C&#7852;P NH&#7852;T 
												TRẢ SÁCH</b></td>
											</tr>
											<tr>
												<td colspan="2">
												<p align="center">
												<img border="0" src="../images/line.gif" width="175" height="5"></td>
											</tr>
											<tr>
												<td colspan="2">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
											</tr>
											<tr>
												<td width="39%" align="right">
												&nbsp;</td>
												<td width="61%">&nbsp;</td>
											</tr>
											<form method="POST" action="admin_returnbook.asp" name="frmBreach">
												</form>
											</tr>
											<tr>
												<td width="39%" align="right">
												<p style="margin-right: 4px">&nbsp;</td>
												<td width="61%">&nbsp;</td>
											</tr>
											<tr>
												<td width="39%" align="right">
												<p style="margin-right: 4px">&nbsp;</td>
												<td width="61%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="131">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							</table>
					</div>
					</td>
				</tr>
				<tr>
					<td colspan="2"><!--#INCLUDE FILE="admin_footer.asp" --></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
</table>

</div>

</body>
</html>
<% End If%>