<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../include/inc_hexpass.asp" -->
<%
	If Session("Mod")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
%>
<%
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM USER WHERE USERNAME='" &Session("Username") & "'"
	rsEdit.CursorType = 2
	rsEdit.LockType = 3
	rsEdit.Open strSQL, Conn
	
	Set rsGroup = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM USERGROUP WHERE ID=" & rsEdit("GROUP_ID")
	rsGroup.Open strSQL, Conn
	
	txtCategory	= Request.Form("category")	
	If txtCategory = "user" Then				
		txtUsername	= Request.Form("txtUsername")
		txtOldPassword	= Request.Form("txtOldPassword")
		
		txtNewPassword	= Request.Form("txtNewPassword")			
		If txtNewPassword <> Request.Form("txtNewPasswordAgain") Then
			Response.Redirect("admin_error.asp?type=3")
		End If	
			
		If txtNewPassword = "" Then
			txtNewPassword = txtOldPassword
		Else
			txtNewPassword = HashEncode(txtNewPassword)	
		End If	
			
		txtFullname	= Request.Form("txtFullname")
		If txtFullname = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtBirthday = Request.Form("txtDay") & "/" & Request.Form("txtMonth") & "/" & Request.Form("txtYear")
		
		strSQL = "UPDATE USER SET "
		strSQL = strSQL & "PASSWORD="&CheckString(txtNewPassword,",")		
		strSQL = strSQL & "FULLNAME="&CheckString(txtFullname,",")		
		strSQL = strSQL & "BIRTHDAY="&CheckString(txtBirthday,"")
		strSQL = strSQL & "WHERE USERNAME='" & txtUsername & "'"

		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_done.asp?page=admin_default.asp")
	Else
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
								<td colspan="4" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px">
								<b>&nbsp; <font color="#FF0000" size="2">THAY ĐỔI THÔNG TIN</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew" action="admin_changepass.asp">
							<tr>
								<td width="79">&nbsp;</td>
								<td width="137">
								<p style="margin-top: 3px; margin-bottom: 3px"><b>Tên truy cập</b></td>
								<td width="348">&nbsp;<b><%=rsEdit("USERNAME")%></b></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="79">
								&nbsp;</td>
								<td width="137">
								<b>Mật khẩu mới</b></td>
								<td width="348">
								<input type="password" name="txtNewPassword" size="30" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="79">
								&nbsp;</td>
								<td width="137">
								<b>Nhập lại mật khẩu mới</b></td>
								<td width="348">
								<input type="password" name="txtNewPasswordAgain" size="30" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="79">
								&nbsp;</td>
								<td width="137">
								<p style="margin-top: 3px; margin-bottom: 3px">
								<b>Nhóm người dùng</b></td>
								<td width="348">&nbsp;<b><%=rsGroup("NAME")%></b></td>
								<td width="9">&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="79">
								&nbsp;</td>
								<td width="137">
								<b>Họ và tên</b></td>
								<td width="348">
								<input type="text" name="txtFullname" size="30" class="input_text" value="<%=rsEdit("FULLNAME")%>"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="79"></td>
								<td height="14" width="137"><b>Ngày sinh</b></td>
								<td height="14" width="348">
								<select size="1" name="txtDay" class="input_text">
								<%
									Call ListNumber(01,31,Day(rsEdit("BIRTHDAY")))
								%>
								</select><select size="1" name="txtMonth" class="input_text">
								<%
									Call ListNumber(01,12,Month(rsEdit("BIRTHDAY")))
								%>
								</select><select size="1" name="txtYear" class="input_text">
								<%
									Call ListNumber(1945,2006,Year(rsEdit("BIRTHDAY")))
								%>
								</select></td>
								<td height="14" width="9">
								</td>
							</tr>
							<tr>
								<td width="573" colspan="4" height="20">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="79">&nbsp;</td>
								<td width="485" colspan="2">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
									<tr>
										<td width="167">&nbsp;</td>
										<td>
								<p align="left">
								<input type="submit" value="Cập nhật" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
									</tr>
								</table>
								</td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="user">
							<input type="hidden" name="txtOldPassword" value="<%=rsEdit("PASSWORD")%>">
							<input type="hidden" name="txtUsername" value="<%=rsEdit("USERNAME")%>">
							</form>
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
<% End If %>