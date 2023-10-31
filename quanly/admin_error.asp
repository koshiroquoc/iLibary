<% Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_parameter.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Refresh" content="6; URL=JavaScript:history.go(-1)">   <!--6 la so giay cho-->
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
								<b>&nbsp;</b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;<font color="#FF0000" size="2">CẢNH BÁO LỖI</font></b></td>
							</tr>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td>
							<p align="center">
							<img border="0" src="../images/iconerror.gif" width="46" height="44"></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td width="150">&nbsp;</td>
							<td><p align="center"><b><font color="#FF0000">
							<% If Request.QueryString("type") = 1 Then %>
								<%=strNoValue%>								
							<% End If %>
							<% If Request.QueryString("type") = 2 Then %>
								<%=strExistValue%>								
							<% End If %>
							<% If Request.QueryString("type") = 3 Then %>
								<%=strNotEqual%>								
							<% End If %>	
							<% If Request.QueryString("type") = 4 Then %>
								<%=strExitsUser%>								
							<% End If %>
							<% If Request.QueryString("type") = 5 Then %>
								<%=strNotPower%>								
							<% End If %>
							<% If Request.QueryString("type") = 6 Then %>
								<%=strShortLen%>								
							<% End If %>
							<% If Request.QueryString("type") = 7 Then %>
								<%=strYearValid%>								
							<% End If %>
							<% If Request.QueryString("type") = 8 Then %>
								<%=strExitCard%>								
							<% End If %>	
							<% If Request.QueryString("type") = 9 Then %>
								<%=strExitBook%>								
							<% End If %>
							<% If Request.QueryString("type") = 10 Then %>
								<%=strEndBook%>								
							<% End If %>
							<% If Request.QueryString("type") = 11 Then %>
								<%=strDoneBook%>								
							<% End If %>
							<% If Request.QueryString("type") = 12 Then %>
								<%=strNoBorrow%>								
							<% End If %>
							<% If Request.QueryString("type") = 13 Then %>
								<%=strExisClass%>								
							<% End If %>
							<% If Request.QueryString("type") = 14 Then %>
								<%=strInvalid %>								
							<% End If %>
							</font></b></td>
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
							<a href="JavaScript:history.go(-1)">vào đây</a></b> hoặc 
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