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
						<table border="0" width="614" id="table3" cellspacing="0" cellpadding="0">
							<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<font color="#FF0000" size="2"><b>&nbsp; 
							</b></font>
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<font color="#FF0000" size="2"><b>CẬP NHẬT TRẢ SÁCH</b></font></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="4">&nbsp;</td>
								<td width="565">
								<table border="0" width="100%" id="table16" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table17" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="3">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
											</tr>
											</tr>
											<tr>
												<td width="14%" align="right">
												<p style="margin-right: 4px">&nbsp;</td>
												<td width="67%">
												<table border="1" width="100%" id="table19" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
													<tr>
														<td>
														<table border="0" width="100%" id="table20" cellspacing="0" cellpadding="0">
															<tr>
																<td width="8%">&nbsp;</td>
																<td width="30%">&nbsp;</td>
																<td width="28%">&nbsp;</td>
																<td width="35%">&nbsp;</td>
															</tr>
															<tr>
																<td colspan="4">
												<p align="center"><b>
												<font size="2">CẬP NHẬT 
												TRẢ SÁCH</font></b></td>
															</tr>
															<tr>
																<td colspan="4">
																<p align="center">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
															</tr>
															<tr>
																<td width="8%">&nbsp;</td>
																<td width="30%">&nbsp;</td>
																<td width="28%">&nbsp;</td>
																<td width="35%">&nbsp;</td>
															</tr>
															<tr>
																<form method="POST" name="frmReturnBook">
																<td width="8%">&nbsp;</td>
																<td width="30%">
																<b>
																<font size="2">Nhập mã sách</font></b></td>
																<td width="28%">
												<font size="2">
												<input type="text" name="txtBookID" size="11" class="input_text"></font></td>
																<td width="35%">
												<input type="submit" value="C&#7853;p nh&#7853;t" name="B2" class="input_button"></td>
																	<input type="hidden" name="category" value="returnbook">
																</form>
															</tr>
															<tr>
																<td width="8%">&nbsp;</td>
																<td width="30%">&nbsp;</td>
																<td width="28%">&nbsp;</td>
																<td width="35%">&nbsp;</td>
															</tr>
															<tr>
																<td width="101%" colspan="4">
																<p align="center">
																<font size="2">Mã sách được lấy 
																từ sách mà độc 
																giả mang đến trả</font></td>
															</tr>
															<tr>
																<td width="8%">&nbsp;</td>
																<td width="30%">&nbsp;</td>
																<td width="28%">&nbsp;</td>
																<td width="35%">&nbsp;</td>
															</tr>
														</table>
														</td>
													</tr>
												</table>
												</td>
												<td width="19%">&nbsp;</td>
											</tr>
											<tr>
												<td width="14%" align="right">
												<p style="margin-right: 4px">&nbsp;</td>
												<td width="67%">&nbsp;</td>
												<td width="19%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="45">&nbsp;</td>
							</tr>
							<%
								txtCategory = Request.Form("category")
								If txtCategory = "returnbook" Then
									txtBookID = Request.Form("txtBookID")
									If txtBookID = "" Then
										Response.Redirect("admin_error.asp?type=1")
									End If
									Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
									strSQL = "SELECT * FROM BORROW WHERE BOOK_ID='" & Trim(txtBookID) & "'"
									rsCheckBook.Open strSQL, Conn,3,1
									If rsCheckBook.Eof Then			
										rsCheckBook.Close
										Set rsCheckBook= Nothing
										Response.Redirect("admin_error.asp?type=9")
									End If
							%>
							<tr>
								<td width="4">&nbsp;</td>
								<td width="565">
								<table border="0" width="100%" id="table12" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table13" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="3">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center"><b>
												<font size="2">NHỮNG THẺ ĐANG MƯỢN MÃ SÁCH: 
												</font>
												<font size="2" color="#0000FF"><%=uCase(rsCheckBook("BOOK_ID"))%></font></b></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td colspan="3">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="1%" align="right">
												<p style="margin-top: 3px; margin-bottom: 3px">&nbsp;</td>
												<td width="98%">
												<table border="1" width="100%" id="table18" bordercolorlight="#F7F7F7" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
													<tr>
														<td width="29" align="center" bgcolor="#CCCCCC">
														<p style="margin-top: 1px; margin-bottom: 1px">
														<b><font size="2">STT</font></b></td>
														<td width="62" align="center" bgcolor="#CCCCCC">
														<p style="margin-top: 1px; margin-bottom: 1px">
														<b><font size="2">Mã thẻ</font></b></td>
														<td width="137" align="center" bgcolor="#CCCCCC">
														<p style="margin-top: 1px; margin-bottom: 1px">
														<b><font size="2">Họ và tên</font></b></td>
														<td width="175" align="center" bgcolor="#CCCCCC">
														<p style="margin-top: 1px; margin-bottom: 1px">
														<b><font size="2">Tên sách</font></b></td>
														<td width="78" align="center" bgcolor="#CCCCCC">
														<p style="margin-top: 1px; margin-bottom: 1px">
														<b><font size="2">Ngày mượn</font></b></td>
														<td align="center" bgcolor="#CCCCCC">&nbsp;</td>
													</tr>
													<%
														Dim iCount
														iCount = 1
														Do While Not rsCheckBook.Eof
														Set rsBook = Server.CreateObject("ADODB.Recordset")
														strSQL = "SELECT * FROM BOOK WHERE BOOK_ID='" & rsCheckBook("BOOK_ID") & "'"
														rsBook.Open strSQL, Conn,3,1
														Set rsUser = Server.CreateObject("ADODB.Recordset")
														strSQL = "SELECT * FROM CARD WHERE CARD_ID='" & rsCheckBook("CARD_ID") & "'"
														rsUser.Open strSQL, Conn,3,1
													%>
													<tr>
														<td width="29">
														<p align="center" style="margin-top: 2px; margin-bottom: 2px">
														<font size="2"><%=iCount%></font></td>
														<td width="62">
														<p align="center" style="margin-top: 2px; margin-bottom: 2px">
														<font size="2"><%=rsCheckBook("CARD_ID")%></font></td>
														<td width="137">
														<p align="left" style="margin-top: 2px;margin-left: 2px; margin-bottom: 2px">
														<font size="2"><%=rsUser("FIRSTNAME") & " " & rsUser("LASTNAME")%></font></td>
														<td width="175">
														<p style="margin:2px 3px; ">
														<font size="2"><%=rsBook("NAME")%></font></td>
														<td width="78">
														<p align="center" style="margin-top: 2px; margin-bottom: 2px">
														<font size="2"><%=NgayVN(rsCheckBook("DATE_INFORM"))%></font></td>
														<td>
														<p align="center" style="margin-top: 2px; margin-bottom: 2px">
														<b>
														<a href="admin_showreturn.asp?id=<%=rsCheckBook("ID")%>">
														<font size="2">Trả sách</font></a></b></td>
													</tr>
													<%
														iCount = iCount + 1
														rsCheckBook.MoveNext
														Loop
														rsBook.Close
														Set rsBook = Nothing
														rsCheckBook.Close
														Set rsCheckBook = Nothing														
													%>
												</table>
												</td>
												<td width="1%">&nbsp;</td>
											</tr>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="45">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<%
								End If
							%>
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