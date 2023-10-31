<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("notice")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>

<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../editor/fckeditor.asp" -->
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
					<td width="187" valign="top" background="../images/bg_menuleft.gif"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="3" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;</b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2">&nbsp;</font><font color="#FF0000" size="2">THỐNG KÊ SÁCH</font></b></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="84">&nbsp;</td>
								<td width="412">
								<table border="1" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center"><b>
												<font size="2">THỐNG KÊ 
												TỔNG QUÁT</font></b></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center" style="margin-bottom: 6px">
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
												<td width="10%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="48%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số sách 
												thư viện hiện có</font></b></td>
												<td width="5%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<font size="2"><%=CountBook()%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<tr>
												<td width="10%" align="right">&nbsp;</td>
												<td width="48%" align="right">
												<p align="left">&nbsp;</td>
												<td width="5%" align="center">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
											<%
												strSQL = "SELECT * FROM CATEGORY_BOOK"
												Set rsCountCate = Conn.Execute(strSQL)
												Do while Not rsCountCate.Eof
											%>
											<tr>
												<td width="10%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="48%" align="left">
												<font color="#003399" size="2"><b>
												<%=rsCountCate("NAME")%></b></font></td>
												<td width="5%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountCateBook(rsCountCate("CATEGORY_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<%
												rsCountCate.MoveNext
												Loop
												rsCountCate.Close
												Set rsCountCate = Nothing
											%>
											<tr>
												<td width="10%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="48%" align="right">
												&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="77">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="84">&nbsp;</td>
								<td width="412">
								<table border="1" width="100%" id="table10" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table11" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center"><b>
												<font size="2">THỐNG KÊ SÁCH CHO MƯỢN</font></b></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center" style="margin-bottom: 6px">
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
												<td width="10%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="48%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số sách đang
												cho mượn</font></b></td>
												<td width="5%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<font size="2"><%=CountBorrow()%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<tr>
												<td width="10%" align="right">&nbsp;</td>
												<td width="48%" align="right">
												<p align="left">&nbsp;</td>
												<td width="5%" align="center">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
											<%
												strSQL = "SELECT * FROM CATEGORY_BOOK"
												Set rsCategory = Conn.Execute(strSQL)
												Do while Not rsCategory.Eof
												strSQL = "SELECT * FROM BORROW WHERE LEFT(BOOK_ID,3)='" & rsCategory("CATEGORY_ID") & "'"
												Set rsBorrow = Conn.Execute(strSQL)
												If Not rsBorrow.Eof Then
											%>
											<tr>
												<td width="10%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="48%" align="left">
												<font color="#003399" size="2"><b>
												<%=rsCategory("NAME")%></b></font></td>
												<td width="5%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountCateBookBorrow(rsCategory("CATEGORY_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<%
												End If
												rsCategory.MoveNext
												Loop
											%>
											<tr>
												<td width="10%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="48%" align="right">
												&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="77">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="84">&nbsp;</td>
								<td width="412">
								<table border="1" width="100%" id="table6" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center"><b>
												<font size="2">THỐNG KÊ 
												MƯỢN - TRẢ</font></b></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center" style="margin-bottom: 6px">
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
												<td width="10%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="48%" align="right">
												<p align="left" style="margin-top: 3px; margin-bottom: 3px"><b>
												<font size="2">Tổng số 
												lượt mượn</font></b></td>
												<td width="5%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<font size="2"><%=CountSumBorrow()%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<tr>
												<td width="10%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="48%" align="right">
												<p align="left" style="margin-top: 3px; margin-bottom: 3px"><b>
												<font size="2">Tổng số 
												lượt trả</font></b></td>
												<td width="5%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="37%"><b><font color="#0000FF">
												<font size="2"><%=CountSumReturn()%></font></font></b><font size="2">&nbsp;&nbsp;cuốn</font></td>
											</tr>
											<tr>
												<td width="10%" align="right">&nbsp;</td>
												<td width="48%" align="right">
												<p align="left">&nbsp;</td>
												<td width="5%" align="center">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
											<tr>
												<td width="10%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="48%" align="right">
												&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="37%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="77">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">
								<p align="center"><b>
								<a href="#" onclick="JavaScript:openWindowPrint('admin_printcountbook.asp')">
								<font color="#FF0000">In thống kê</font></a></b></td>
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
