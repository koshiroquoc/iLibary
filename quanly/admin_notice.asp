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
								<td colspan="7" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<font color="#FF0000" size="2">
								<b>&nbsp;QUẢN LÝ THÔNG BÁO</b></font></td>
							</tr>
							<tr>
								<td colspan="7">&nbsp;</td>
							</tr>
							<tr>
								<td width="78" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="9" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="192" height="33">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="4" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="175" height="33">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="63" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="9">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="192">
								<p style="margin-top: 0; margin-bottom: 3px" align="center">
								<a href="admin_addnotice.asp">
								<img border="0" src="../images/addnotice.gif" width="40" height="39"></a></td>
								<td width="4">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="175">
								<p style="margin-top: 0; margin-bottom: 3px" align="center">
								<a href="admin_listnotice.asp">
								<img border="0" src="../images/list.gif" width="40" height="39"></a></td>
								<td width="10">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="9">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="192">
								<p align="center" style="margin-top: 3px"><b>
								<font size="2">Thông báo mới</font></b></td>
								<td width="4">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="175">
								<p align="center" style="margin-top: 3px"><b>
								<font size="2">Liệt kê - Sửa đổi</font></b></td>
								<td width="10">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-top: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="192">&nbsp;</td>
								<td width="4">&nbsp;</td>
								<td width="175">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="192">&nbsp;</td>
								<td width="4">&nbsp;</td>
								<td width="175">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="7">&nbsp;</td>
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
