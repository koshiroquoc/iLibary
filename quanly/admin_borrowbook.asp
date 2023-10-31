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
	txtCategory	= Request.Form("category")	
	If txtCategory = "borrowbook" Then
		txtCardID = Request.Form("txtCardID")
		If txtCardID = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtBookID = Request.Form("txtBookID")
		If txtBookID = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		Set rsCheckCard = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT CARD_ID,CLASS_ID FROM CARD WHERE CARD_ID='" & Trim(txtCardID) & "'"
		rsCheckCard.Open strSQL, Conn,3,1
		If rsCheckCard.Eof Then			
			rsCheckCard.Close
			Set rsCheckCard = Nothing
			Response.Redirect("admin_error.asp?type=8")
		Else
			txtClassID = rsCheckCard("CLASS_ID")
			rsCheckCard.Close
			Set rsCheckCard = Nothing
		End If
		
		Set rsBreach = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT CARD_ID FROM BREACH WHERE CARD_ID='" & Trim(txtCardID) & "'"
		rsBreach.Open strSQL, Conn,3,1
		If Not rsBreach.Eof Then			
			rsBreach.Close
			Set rsBreach = Nothing
			Response.Redirect("admin_breach.asp")
		Else
			rsBreach.Close
			Set rsBreach = Nothing				
		End If
		
		Set rsDoneCheck = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT CARD_ID FROM BORROW WHERE CARD_ID='" & Trim(txtCardID) & "'"
		rsDoneCheck.Open strSQL, Conn,3,1
		If Not rsDoneCheck.Eof Then			
			rsDoneCheck.Close
			Set rsDoneCheck = Nothing
			Response.Redirect("admin_error.asp?type=11")
		Else
			rsDoneCheck.Close
			Set rsDoneCheck = Nothing			
		End If
		
		Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT BOOK_ID, AMOUNT FROM BOOK WHERE BOOK_ID='" & Trim(txtBookID) & "'"
		rsCheckBook.Open strSQL, Conn,3,1
		If rsCheckBook.Eof Then			
			rsCheckBook.Close
			Set rsCheckBook = Nothing
			Response.Redirect("admin_error.asp?type=9")
		Else
			If rsCheckBook("AMOUNT") = CheckCountBorrow(txtBookID) Then
				rsCheckBook.Close
				Set rsCheckBook = Nothing
				Response.Redirect("admin_error.asp?type=10")
			End If				
		End If
		
		strSQL = "SELECT * FROM BORROW Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO BORROW(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
		strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
		strSQL = strSQL & CheckString(Now(),")")
		Conn.Execute(strSQL)
		
		strSQL = "SELECT * FROM TEMP_BORROW Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO TEMP_BORROW(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
		strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
		strSQL = strSQL & CheckString(Now(),")")
		Conn.Execute(strSQL)
		Conn.Close
		Set Conn = Nothing
		Response.Redirect("admin_doneborrow.asp")
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
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2">&nbsp; <font color="#FF0000">CẬP NHẬT MƯỢN SÁCH</font></font></b></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="83">&nbsp;</td>
								<td width="388">
								<table border="1" width="100%" id="table12" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="113%" id="table13" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="3">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center"><b>
												<font size="2">C&#7852;P NH&#7852;T M&#431;&#7906;N SÁCH</font></b></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="40%" align="right">
												&nbsp;</td>
												<td width="60%" colspan="2">&nbsp;</td>
											</tr>
											<form method="POST" action="admin_borrowbook.asp" name="frmBorrow">
											<tr>
												<td width="40%" align="right">
												<p style="margin-right: 4px"><b>
												<font size="2">Nh&#7853;p mã th&#7867; 
												</font> </b></td>
												<td width="22%">
												<font size="2">
												<input type="text" name="txtCardID" size="8" class="input_text"></font></td>
												<td width="39%">
												&nbsp;</td>
												</tr>
											<tr>
												<td width="40%" align="right">
												<p style="margin-right: 4px"><b>
												<font size="2">Nh&#7853;p mã sách 
												</font> </b></td>
												<td width="22%">
												<font size="2">
												<input type="text" name="txtBookID" size="8" class="input_text"></font></td>
												<td width="39%">
												<font size="2">
												<input type="submit" value="C&#7853;p nh&#7853;t" name="B1" class="input_button"></font></td>
												<input type="hidden" name="category" value="borrowbook">
												</form>
											</tr>
											<tr>
												<td width="40%" align="right">
												<p style="margin-right: 4px">&nbsp;</td>
												<td width="60%" colspan="2">&nbsp;</td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="3">
												<p style="margin-right: 4px" align="center">
												<font size="2">Mã thẻ được lấy từ Thẻ thư viện 
												của độc giả.<br>
												Mã sách được cung cấp từ độc 
												giả.</font></td>
											</tr>
											<tr>
												<td width="40%" align="right">
												&nbsp;</td>
												<td width="60%" colspan="2">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="102">&nbsp;</td>
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