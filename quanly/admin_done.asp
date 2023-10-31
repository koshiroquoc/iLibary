<% Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_parameter.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Refresh" content="3; URL=JavaScript:history.go(-1)">   <!--6 la so giay cho-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body topmargin="8" leftmargin="8">

<table border="1" width="984" id="table1" bordercolordark="#808080" cellspacing="0" cellpadding="0" bordercolorlight="#D5F1FF" align="center">
	<tr>
		<td>
		<div align="center">
			<table border="0" width="984" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td colspan="2"><!--#INCLUDE FILE="admin_header.asp" --></td>
				</tr>
				<tr>
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width="797" height="350" valign="top">
					<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<font color="#FF0000" size="2"><b>&nbsp; THÔNG 
							TIN HỆ THỐNG</b></font></td>
							</tr>
						<tr>
							<td colspan="3">&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>
							<p align="center">
							&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td width="150">&nbsp;</td>
							<td><p align="center">
							<font color="#0000FF" size="2"><b>Dữ 
							liệu đã được cập nhật thành công!</b></font></td>
							<td width="150">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="3">&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>
							<p align="center"><font size="2">Bấm
							<b>
							<a href="<%=Request.QueryString("page")%>">vào đây</a></b> hoặc 
							chờ trong vài giây hệ thống sẽ tự động quay trở lại.</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</table>
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

</body>

</html>