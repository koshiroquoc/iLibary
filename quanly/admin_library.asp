<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("library")= "" Then
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
<!-- #INCLUDE FILE="../include/inc_js.asp" -->
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
						<table border="0" width="658" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="3" height="48">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>
								<font style="font-size: 11pt" color="#FF0000">QUẢN LÝ THƯ VIỆN</font></b></td>
							</tr>
							<tr>
								<td width="12">&nbsp;</td>
								<td width="594">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="106%" cellspacing="0" cellpadding="0">
											<tr>
												<td width="10">&nbsp;</td>
												<td width="145">&nbsp;</td>
												<td width="9">&nbsp;</td>
												<td width="130">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="15">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="112">&nbsp;</td>
												<td width="10">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_borrowbook.asp" onMouseOver="JavaScript:handleOver('borrow','borrow1');return true;" onMouseOut="JavaScript:handleOut('borrow','borrow');return true;">
												<font style="font-size: 11pt">
												<img border="0" name=borrow src="../images/borrow.gif" width="48" height="47"></font></a></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_returnbook.asp" onMouseOver="JavaScript:handleOver('return','return1');return true;" onMouseOut="JavaScript:handleOut('return','return');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="return" src="../images/return.gif" width="48" height="47"></font></a></td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_addnotice.asp" onMouseOver="JavaScript:handleOver('notice','notice1');return true;" onMouseOut="JavaScript:handleOut('notice','notice');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="notice" src="../images/notice.gif" width="48" height="49"></font></a></td>
												<td>&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_addschedule.asp" onMouseOver="JavaScript:handleOver('schedule','schedule1');return true;" onMouseOut="JavaScript:handleOut('schedule','schedule');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="schedule" src="../images/schedule.gif" width="48" height="47"></font></a></td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">
												<p align="center">
												<font style="font-size: 11pt">Cập nhật mượn 
												sách</font></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Cập nhật trả 
												sách</font></td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Cập nhật thông 
												báo</font></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Cập nhật lịch 
												trực</font></td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="22"></td>
												<td width="145" height="22"></td>
												<td width="9" height="22"></td>
												<td height="22"></td>
												<td width="12" height="22"></td>
												<td height="22"></td>
												<td width="13" height="22"></td>
												<td height="22"></td>
												<td height="22"></td>
												<td height="22"></td>
												<td width="4" height="22"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="145">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_adddoc.asp" onMouseOver="JavaScript:handleOver('document','document1');return true;" onMouseOut="JavaScript:handleOut('document','document');return true;">
												<font style="font-size: 11pt">
												<img border="0" name ="document" src="../images/document.gif" width="48" height="47"></font></a></td>
												<td width="9">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_listregister.asp" onMouseOver="JavaScript:handleOver('register','register1');return true;" onMouseOut="JavaScript:handleOut('register','register');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="register" src="../images/register.gif" width="48" height="47"></font></a></td>
												<td width="12">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_countbook.asp" onMouseOver="JavaScript:handleOver('countbook','countbook1');return true;" onMouseOut="JavaScript:handleOut('countbook','countbook');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="countbook" src="../images/countbook.gif" width="48" height="47"></font></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_countcard.asp" onMouseOver="JavaScript:handleOver('countcard','countcard1');return true;" onMouseOut="JavaScript:handleOut('countcard','countcard');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="countcard" src="../images/countcard.gif" width="48" height="47"></font></a></td>
												<td width="4">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">
												<p align="center">
												<font style="font-size: 11pt">Cập nhật tài 
												liệu</font></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Độc giả đăng 
												ký</font></td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Thống kê sách</font></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Thống kê thẻ</font></td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="25"></td>
												<td width="145" height="25"></td>
												<td width="9" height="25"></td>
												<td height="25"></td>
												<td width="12" height="25"></td>
												<td height="25"></td>
												<td width="13" height="25"></td>
												<td height="25"></td>
												<td height="25"></td>
												<td height="25"></td>
												<td width="4" height="25"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="145">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_libraring.asp" onMouseOver="JavaScript:handleOver('library','library1');return true;" onMouseOut="JavaScript:handleOut('library','library');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="library" src="../images/library.gif" width="48" height="47"></font></a></td>
												<td width="9">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listnotice.asp" onMouseOver="JavaScript:handleOver('noticemana','noticemana1');return true;" onMouseOut="JavaScript:handleOut('noticemana','noticemana');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="noticemana" src="../images/noticemana.gif" width="48" height="47"></font></a></td>
												<td width="12">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listdoc.asp" onMouseOver="JavaScript:handleOver('documentmana','documentmana1');return true;" onMouseOut="JavaScript:handleOut('documentmana','documentmana');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="documentmana" src="../images/documentmana.gif" width="48" height="47"></font></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listschedule.asp" onMouseOver="JavaScript:handleOver('schedulemana','schedulemana1');return true;" onMouseOut="JavaScript:handleOut('schedulemana','schedulemana');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="schedulemana" src="../images/schedulemana.gif" width="48" height="47"></font></a></td>
												<td width="4">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">
												<p align="center">
												<font style="font-size: 11pt">Quản lý mượn 
												trả</font></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Quản lý thông 
												báo</font></td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Quản lý tài 
												liệu</font></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Quản lý lịch 
												trực</font></td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="25"></td>
												<td width="145" height="25"></td>
												<td width="9" height="25"></td>
												<td height="25"></td>
												<td width="12" height="25"></td>
												<td height="25"></td>
												<td width="13" height="25"></td>
												<td height="25"></td>
												<td height="25"></td>
												<td height="25"></td>
												<td width="4" height="25"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="145">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="9">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listbook.asp" onMouseOver="JavaScript:handleOver('book1','book11');return true;" onMouseOut="JavaScript:handleOut('book1','book1');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="book1" src="../images/book1.gif" width="48" height="47"></font></a></td>
												<td width="12">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_listcard.asp" onMouseOver="JavaScript:handleOver('cardmana','cardmana1');return true;" onMouseOut="JavaScript:handleOut('cardmana','cardmana');return true;">
												<font style="font-size: 11pt">
												<img border="0" name="cardmana" src="../images/cardmana.gif" width="48" height="47"></font></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="4">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">
												<p align="center">&nbsp;</td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Quản lý sách</font></td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center">
												<font style="font-size: 11pt">Quản lý thẻ</font></td>
												<td>
												<p align="center">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="4">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="145">&nbsp;</td>
												<td width="9">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="12">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
												<td>&nbsp;</td>
												<td width="4">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="52">&nbsp;</td>
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