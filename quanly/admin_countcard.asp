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
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="3" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;</b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;<font color="#FF0000" size="2">THỐNG KÊ THẺ THƯ VIỆN</font></b></td>
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
												<img border="0" src="../images/spacer.gif" width="1" height="10"></td>
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
												<td width="23%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="46%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số thẻ 
												thư viện hiện có</font></b></td>
												<td width="1%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="31%"><b><font color="#0000FF">
												<font size="2"><%=CountCard()%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p style="margin-top: 10px; margin-bottom: 10px" align="center">
												<font size="2">
												<img border="0" src="../images/line.gif" width="110" height="5"></font></td>
											</tr>
											<%
												strSQL = "SELECT DISTINCT CLASS_ID FROM CARD"
												Set rsCountCate = Conn.Execute(strSQL)
												Do while Not rsCountCate.Eof
											%>
											<tr>
												<td width="23%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="46%" align="left"><b>
												<font size="2">Lớp:&nbsp;&nbsp;</font><font size="2" color="#003399"><%= rsCountCate("CLASS_ID")%></font></b></td>
												<td width="1%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="31%"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountClassCard(rsCountCate("CLASS_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<%
												rsCountCate.MoveNext
												Loop
												rsCountCate.Close
												Set rsCountCate = Nothing
											%>
											<tr>
												<td width="23%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="46%" align="right">
												&nbsp;</td>
												<td width="1%">&nbsp;</td>
												<td width="31%">&nbsp;</td>
											</tr>
											<tr>
												<td width="101%" align="right" colspan="4">
												<table border="0" width="100%" id="table12" cellspacing="0" cellpadding="0">
													<tr>
														<td>
														<p style="margin-top: 8px">&nbsp;</td>
														<td width="14">
														<p align="center" style="margin-top: 8px">
														<font size="2">
														<img border="0" src="../images/close.gif" width="14" height="9"></font></td>
														<td width="115">
												<p align="center" style="margin-top: 8px"><b>
												<a href="admin_listcard.asp">
												<font size="2">Danh sách thẻ</font></a></b></td>
														<td width="126">
														<p style="margin-top: 8px">&nbsp;</td>
													</tr>
												</table>
												</td>
												</tr>
											<tr>
												<td width="23%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="46%" align="right">
												&nbsp;</td>
												<td width="1%">&nbsp;</td>
												<td width="31%">&nbsp;</td>
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
												<td colspan="6">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center"><b>
												<font size="2">THỐNG KÊ 
												THẺ MƯỢN QUÁ HẠN</font></b></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center" style="margin-bottom: 6px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="53%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số thẻ 
												đang mượn quá hạn</font></b></td>
												<td width="2%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="6%"><b><font color="#0000FF">
												<font size="2"><%=CountBreaching()%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
												<td width="4%">
												<p align="center">
												<a href="JavaScript:openWindowPrint('admin_print.asp?typeprint=breaching&class=All');">
												<font size="2">
												<img border="0" src="../images/print.gif" width="15" height="11" alt="In danh sách"></font></a></td>
												<td width="16%">&nbsp;</td>
											</tr>
											<tr>
												<td width="19%" align="right">&nbsp;</td>
												<td width="53%" align="right">
												<p align="left">&nbsp;</td>
												<td width="2%" align="center">&nbsp;</td>
												<td width="6%">&nbsp;</td>
												<td width="4%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
											</tr>
											<%
												strSQL = "SELECT DISTINCT CLASS_ID FROM CARD"
												Set rsCategory = Conn.Execute(strSQL)
												Do while Not rsCategory.Eof
												If CountBreachingClass(rsCategory("CLASS_ID"))>0 Then
											%>
											<tr>
												<td width="19%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="53%" align="left"><b>
												<font size="2">Lớp:</font><font size="2" color="#003399">&nbsp;&nbsp;<%=rsCategory("CLASS_ID")%></font></b></td>
												<td width="2%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="6%"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountBreachingClass(rsCategory("CLASS_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
												<td width="4%">
												<p align="center">
												<a href="JavaScript:openWindowPrint('admin_print.asp?typeprint=breaching&class=<%=rsCategory("CLASS_ID")%>');">
												<font size="2">
												<img border="0" src="../images/print.gif" width="15" height="11" alt="In danh sách"></font></a></td>
												<td width="16%">&nbsp;</td>
											</tr>
											<%
												End If
												rsCategory.MoveNext
												Loop
											%>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="53%" align="right">
												&nbsp;</td>
												<td width="2%">&nbsp;</td>
												<td width="6%">&nbsp;</td>
												<td width="4%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="6">
												<p style="margin-right: 3px" align="center">
												<table border="0" width="100%" id="table12" cellspacing="0" cellpadding="0">
													<tr>
														<td>
														<p style="margin-top: 8px">&nbsp;</td>
														<td width="24">
														<p align="center" style="margin-top: 8px">
														<font size="2">
														<img border="0" src="../images/close.gif" width="14" height="9"></font></td>
														<td width="128">
												<p align="center" style="margin-top: 8px"><b>
												<a href="admin_breaching.asp">
												<font size="2">Danh sách 
												quá hạn</font></a></b></td>
														<td width="124">
														<p style="margin-top: 8px">&nbsp;</td>
													</tr>
												</table>
												</td>
												</tr>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="53%" align="right">
												&nbsp;</td>
												<td width="2%">&nbsp;</td>
												<td width="6%">&nbsp;</td>
												<td width="4%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
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
								<a href="#" onclick="JavaScript:openWindowPrint('admin_printcountcard.asp')">
								<font size="2">In thống kê</font></a></b></td>
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
