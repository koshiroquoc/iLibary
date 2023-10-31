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
						<table border="0" width="797" id="table3" cellspacing="0" cellpadding="0">
							<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<b><font color="#FF0000" size="2"> 
							CẬP NHẬT TRẢ SÁCH</font></b><p style="margin-top: 2px; margin-bottom: 2px" align="center">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<%
								id = Request.QueryString("id")
								If id <> "" Then
									Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
									strSQL = "SELECT * FROM BORROW WHERE ID=" & id
									rsCheckBook.Open strSQL, Conn,3,1
									If Not rsCheckBook.Eof Then
										txtCardID = rsCheckBook("CARD_ID")
										txtBookID = rsCheckBook("BOOK_ID")
										txtClassID = rsCheckBook("CLASS_ID")
										txtDateBorrow = rsCheckBook("DATE_INFORM")
										txtBreach = "No"
		
										Set rsBookName = Server.CreateObject("ADODB.Recordset")
										strSQL = "SELECT BOOK_ID,NAME FROM BOOK WHERE BOOK_ID='" & Trim(txtBookID) & "'"
										rsBookName.Open strSQL, Conn,3,1
										txtBookName = rsBookName("NAME")
										If ( Now() > rsCheckBook("DATE_INFORM") + 7) Then
		
											' Add record in Breach table
											strSQL = "SELECT * FROM BREACH Order by ID Desc"		
											txtID = GetID(strSQL,Conn)
											strSQL = "INSERT INTO BREACH(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
											strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
											strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
											strSQL = strSQL & CheckString(Now(),")")
											Conn.EXECUTE(strSQL)
											txtBreach = "Yes"
										End If
		
										' Add record in TEMP_RETURN table
										strSQL = "SELECT * FROM TEMP_RETURN Order by ID Desc"		
										txtID = GetID(strSQL,Conn)
										strSQL = "INSERT INTO TEMP_RETURN(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
										strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
										strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
										strSQL = strSQL & CheckString(Now(),")")
										Conn.EXECUTE(strSQL)
		
										' Delete record in Borrow table
										strSQL = "DELETE * FROM BORROW WHERE CARD_ID='" & txtCardID & "'"
										Conn.EXECUTE(strSQL)							
									End If
									Conn.Close
									Set Conn = Nothing
								End If					
							%>
							<tr>
								<td width="30">&nbsp;</td>
								<td width="515">
								<table border="1" width="100%" id="table12" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table13" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="10"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center"><b>
												<font size="2">THÔNG TIN TRẢ SÁCH</font></b></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="15%">
												<p style="margin-top: 3px; margin-bottom: 3px">
												<b><font size="2">Mã sách</font></b></td>
												<td width="2%" align="center">
												<p align="center"><b>
												<font size="2">:</font></b></td>
												<td width="78%"><font size="2">&nbsp;</font><font color="#008000"><b><font size="2"><%=uCase(txtBookID)%></font></b></font></td>
											</tr>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="15%">
												<p style="margin-top: 3px; margin-bottom: 3px">
												<b><font size="2">Tên sách</font></b></td>
												<td width="2%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2">&nbsp;</font><font color="#008000"><b><font size="2"><%=txtBookName%></font></b></font></td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 4px; margin-top:3px; margin-bottom:3px">&nbsp;</td>
												<td width="15%">
												<p style="margin-top: 3px; margin-bottom: 3px">
												<b><font size="2">Ngày mượn</font></b></td>
												<td width="2%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2">&nbsp;</font><font color="#008000"><b><font size="2"><%=NgayVN(txtDateBorrow)%></font></b></font></td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="15%">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="2%" align="center">
												&nbsp;</td>
												<td width="78%">&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="15%">
												<p style="margin-top: 3px; margin-bottom: 3px">
												<b><font size="2">Mã thẻ</font></b></td>
												<td width="2%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2">&nbsp;</font><font color="#008000"><b><font size="2"><%=uCase(txtCardID)%></font></b></font></td>
											</tr>
											<%
												If txtBreach = "Yes" Then
											%>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p align="center" style="margin-top: 5px; margin-bottom: 5px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="405" height="5"></font></td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p align="center" style="margin-top: 5px; margin-bottom: 5px">
												<font color="#FF0000" size="2"><b>Trường 
												hợp này đã vi phạm quy chế mượn 
												sách. </b></font></td>
											</tr>
											<%
												End If
											%>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p align="center" style="margin-top: 5px; margin-bottom: 5px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="405" height="5"></font></td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p align="center"><b>
												<font size="2">Đã cập nhật 
												trả sách thành công. <br>
												Kích vào
												<a href="admin_returnbook.asp">
												&nbsp;Quay lại</a> để thực hiện 
												tiếp công việc.</font></b></td>
											</tr>
											<tr>
												<td width="5%" align="right">
												&nbsp;</td>
												<td width="15%">&nbsp;</td>
												<td width="2%" align="center">&nbsp;</td>
												<td width="78%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="28">&nbsp;</td>
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