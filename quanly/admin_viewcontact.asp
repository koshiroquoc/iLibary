<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("contact")= False Then
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
<%
	id = Request.QueryString("id")
	strSQL = "SELECT * FROM CONTACT WHERE ID =" & id
	Set rsContact = Server.CreateObject("ADODB.Recordset")
	rsContact.Open strSQL,Conn,3,1
%>
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
						<table border="0" width="797" id="table3" cellspacing="0" cellpadding="0">
							<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>&nbsp; 
							<font color="#FF0000" size="2">NỘI DUNG GÓP Ý</font></b></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="23">&nbsp;</td>
								<td width="530">
								<table border="1" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="5">
												<p align="center">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
											</tr>
											<tr>
												<td colspan="5">
												<p align="center"><b>
												<font size="2">NỘI DUNG 
												GÓP Ý</font></b></td>
											</tr>
											<tr>
												<td colspan="5">
												<p align="center" style="margin-bottom: 6px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="5">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="12%" align="right">
												<p align="left" style="margin-top: 2px; margin-bottom: 2px">
												<font size="2">Người gửi</font></td>
												<td width="2%" align="left">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2"><%=rsContact("FULLNAME")%></font></td>
												<td width="4%">&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="12%" align="right">
												<p align="left" style="margin-top: 2px; margin-bottom: 2px">
												<font size="2">Email</font></td>
												<td width="2%" align="left">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2"><%=rsContact("EMAIL")%></font></td>
												<td width="4%">&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="12%" align="right">
												<p align="left" style="margin-top: 2px; margin-bottom: 2px">
												<font size="2">Ngày gửi</font></td>
												<td width="2%" align="left">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2"><%=NgayVN(rsContact("DATE_INFORM"))%></font></td>
												<td width="4%">&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="12%" align="right">
												<p align="left" style="margin-top: 2px; margin-bottom: 2px">
												<font size="2">Tiêu đề</font></td>
												<td width="2%" align="left">
												<b><font size="2">:</font></b></td>
												<td width="78%"><font size="2"><%=rsContact("TITLE")%></font></td>
												<td width="4%">&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="12%" align="right">
												<p align="left" style="margin-top: 2px; margin-bottom: 2px">
												<font size="2">Nội dung</font></td>
												<td width="2%" align="left">
												<b><font size="2">:</font></b></td>
												<td width="78%">
												<p align="justify">
												<font size="2"><%=rsContact("CONTENT")%></font></td>
												<td width="4%">
												&nbsp;</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="12%" align="right">
												&nbsp;</td>
												<td width="2%">&nbsp;</td>
												<td width="82%" colspan="2">&nbsp;</td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="5">
												<table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
													<tr>
														<td>&nbsp;</td>
														<td width="20">
								<font size="2">
								<img border="0" src="../images/left.gif" width="22" height="20"></font></td>
														<td width="81"><b>
														<a href="JavaScript:history.back();">
														<font size="2">Quay lại</font></a></b></td>
														<td width="23">
														<p align="center">
														<font size="2">
														<img border="0" src="../images/ed_delete.gif" width="18" height="18"></font></td>
														<td width="231">
														<font size="2">&nbsp;</font><b><a href="admin_delete.asp?category=contact&id=<%=rsContact("ID")%>"><font size="2"> 
														Xóa</font></a></b></td>
													</tr>
												</table>
												</td>
											</tr>
											<tr>
												<td width="5%" align="right">
												&nbsp;</td>
												<td width="95%" align="right" colspan="4">
												&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="20">&nbsp;</td>
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
