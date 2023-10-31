<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
'<%
'	If Session("Username")= "" Then
'		Response.Redirect("admin_login.asp")
'	End If	
'%>
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
								<font color="#FF0000" size="2">
								<b>&nbsp;QUẢN LÝ NGƯỜI DÙNG</b></font></td>
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
												<td width="90" height="35">&nbsp;</td>
												<td width="11" height="35">&nbsp;</td>
												<td width="138" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="28" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="112" height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="90">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												&nbsp;</td>
												<td width="11">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_adduser.asp">
												<font size="2">
												<img border="0" src="../images/user.gif" width="40" height="39"></font></a></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_listuser.asp">
												<font size="2">
												<img border="0" src="../images/list.gif" width="40" height="39"></font></a></td>
												<td>&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="90">
												&nbsp;</td>
												<td width="11">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Tạo người 
												dùng mới</font></b></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Liệt kê - 
												Sửa đổi</font></b></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="17"></td>
												<td width="90" height="17"></td>
												<td width="11" height="17"></td>
												<td height="17"></td>
												<td width="8" height="17"></td>
												<td height="17"></td>
												<td width="13" height="17"></td>
												<td height="17"></td>
												<td height="17"></td>
												<td height="17"></td>
												<td width="10" height="17"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="90">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												&nbsp;</td>
												<td width="11">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_setuser.asp">
												<font size="2">
												<img border="0" src="../images/userpre.gif" width="40" height="39"></font></a></td>
												<td width="8">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_setgroup.asp">
												<font size="2">
												<img border="0" src="../images/grouppre.gif" width="40" height="39"></font></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												&nbsp;</td>
												<td width="10">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="90">
												<p align="center">&nbsp;</td>
												<td width="11">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Phân quyền<br>
												cho người dùng</font></b></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center"><b>
												<font size="2">Phân quyền<br>
												nhóm người dùng</font></b></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="90" height="10"></td>
												<td width="11" height="10"></td>
												<td height="10"></td>
												<td width="8" height="10"></td>
												<td height="10"></td>
												<td width="13" height="10"></td>
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