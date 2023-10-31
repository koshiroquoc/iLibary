<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_parameter.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Refresh" content="3; URL=JavaScript:history.go(-1)">   <!--6 la so giay cho-->
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="3">
		<img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td width="12">&nbsp;</td>
		<td width="411">
		<table border="1" width="100%" id="table2" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
			<tr>
				<td>
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="62">&nbsp;</td>
						<td width="277">
						<p align="center" style="margin-top: 8px; margin-bottom: 6px">
						&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td width="62" height="54">&nbsp;</td>
						<td width="277" height="54">
						<p align="center">
							<img border="0" src="../images/iconerror.gif" width="46" height="44"></td>
						<td height="54">&nbsp;</td>
					</tr>
					<tr>
						<td width="62">&nbsp;</td>
						<td width="277"><p align="center"><b><font color="#FF0000">
							<% If Request.QueryString("type") = 1 Then %>
								<%=strKeyBlank%>								
							<% End If %>
							<% If Request.QueryString("type") = 2 Then %>
								<%=strExistValue%>								
							<% End If %>
						</font></b>
						</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td colspan="3" height="4"></td>
					</tr>
					<tr>
						<td width="62">&nbsp;</td>
						<td width="277">
						<p align="center">Bấm
							<b>
							<a href="JavaScript:history.go(-1)">vào đây</a></b> hoặc 
							chờ trong vài giây hệ thống sẽ tự động quay trở lại.</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td width="62">&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td width="62">&nbsp;</td>
						<td width="277">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td width="62">&nbsp;</td>
						<td width="277">
						&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
		<td width="12">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
</table>

</div>
</body>

</html>