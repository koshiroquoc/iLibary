<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
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
								<b><font size="2" color="#FF0000">&nbsp;</font></b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								&nbsp;<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="3" color="#FF0000">QUẢN LÝ MƯỢN - 
								TRẢ SÁCH</font></b></td>
							</tr>
							<tr>
								<td width="12">&nbsp;</td>
								<td width="548">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" cellspacing="0" cellpadding="0">
											<tr>
												<td width="10" height="35">&nbsp;</td>
												<td width="159" height="35">&nbsp;</td>
												<td width="35" height="35">&nbsp;</td>
												<td width="115" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="7" height="35">&nbsp;</td>
												<td width="162" height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="159">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_borrowbook.asp">
												<font size="2">
												<img border="0" name=borrow src="../images/borrow.gif" width="48" height="47"></font></a></td>
												<td width="35">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_returnbook.asp">
												<font size="2">
												<img border="0" src="../images/return.gif" width="48" height="47"></font></a></td>
												<td width="17">&nbsp;</td>
												<td>&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_listregister.asp" onmouseover>
												<font size="2">
												<img border="0" src="../images/register.gif" width="48" height="47"></font></a></td>
												<td>&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="159">
												<p align="center"><b>
												<font size="2">Cập nhật mượn 
												sách</font></b></td>
												<td width="35">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Cập nhật trả 
												sách</font></b></td>
												<td width="17">&nbsp;</td>
												<td>&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Danh sách đăng 
												ký mượn</font></b></td>
												<td>&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10"></td>
												<td width="35" height="10"></td>
												<td height="10"></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="159">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_countbook.asp">
												<font size="2">
												<img border="0" src="../images/countbook.gif" width="48" height="47"></font></a><p align="center" style="margin-top: 0; margin-bottom: 5px">
												&nbsp;</td>
												<td width="35">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listbreach.asp">
												<font size="2">
												<img border="0" src="../images/breach.gif" width="48" height="47"></font></a></td>
												<td width="17">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_breaching.asp">
												<font size="2">
												<img border="0" src="../images/breaching.gif" width="40" height="39"></font></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="10">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="159">
												<p align="center"><b>
												<font size="2">Thống kê sách</font></b></td>
												<td width="35">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Độc giả vi 
												phạm</font></b></td>
												<td width="17">&nbsp;</td>
												<td>&nbsp;</td>
												<td>
												<p align="center">
												<font size="2"><b>Danh sách</b></font><b><font size="2"> mượn quá 
												hạn</font></b></td>
												<td>&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10"></td>
												<td width="35" height="10"></td>
												<td height="10"></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10">
												<p align="center">
												<a href="admin_listborrow.asp">
												<img border="0" src="../images/card.gif" width="48" height="47"></a></td>
												<td width="35" height="10"></td>
												<td height="10">
												<p align="center">
												<a href="admin_listreturn.asp">
												<img border="0" src="../images/card_new.gif" width="47" height="46"></a></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10">
												<p align="center"><b>
												<font size="2">Danh sách đang 
												mượn</font></b></td>
												<td width="35" height="10"></td>
												<td height="10">
												<p align="center"><b>
												<font size="2">Danh sách trả</font></b></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10"></td>
												<td width="35" height="10"></td>
												<td height="10"></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="159" height="10"></td>
												<td width="35" height="10"></td>
												<td height="10"></td>
												<td width="17" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="13">&nbsp;</td>
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