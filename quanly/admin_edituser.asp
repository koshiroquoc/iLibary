<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If		
	If Session("Admin") = False Then
		Response.Redirect("admin_error.asp?type=5")
	End If
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../include/inc_hexpass.asp" -->
<%
	id	= Request.QueryString("id")	
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM USER WHERE ID="&id
	rsEdit.CursorType = 2
	rsEdit.LockType = 3
	rsEdit.Open strSQL, Conn
	txtUserNameOld = rsEdit("USERNAME")
	
	txtCategory	= Request.Form("category")	
	If txtCategory = "user" Then				
		txtUsername	= Request.Form("txtUsername")
		If txtUsername = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		If Trim(txtUsername) <> txtUserNameOld Then
			Set rsCheck = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM USER WHERE USERNAME='" & Trim(txtUsername) & "'"
			strSQL = strSQL & " AND ID <>" &id
			rsCheck.Open strSQL, Conn,3,1
			If Not rsCheck.Eof Then			
				rsCheck.Close
				Set rsCheck = Nothing
				Response.Redirect("admin_error.asp?type=2")
			End If
		End If
				
		txtPassword	= Request.Form("txtPassword")
		If txtPassword = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		If txtPassword <> Request.Form("txtPasswordAgain") Then
			Response.Redirect("admin_error.asp?type=3")
		End If		
		
		
		txtPassword = HashEncode(txtPassword)

		txtGroup	= Request.Form("txtGroup")

		txtFullname	= Request.Form("txtFullname")
		If txtFullname = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtBirthday = Request.Form("txtDay") & "/" & Request.Form("txtMonth") & "/" & Request.Form("txtYear")
		
		strSQL = "UPDATE USER SET "
		strSQL = strSQL & "USERNAME="&CheckString(txtUsername,",")
		strSQL = strSQL & "PASSWORD="&CheckString(txtPassword,",")		
		strSQL = strSQL & "FULLNAME="&CheckString(txtFullname,",")		
		strSQL = strSQL & "GROUP_ID="&CheckString(txtGroup,",")				
		strSQL = strSQL & "BIRTHDAY="&CheckString(txtBirthday,"")
		strSQL = strSQL & "WHERE ID="& id

		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listuser.asp")
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
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp; <font color="#FF0000" size="2">HIỆU CHỈNH NGƯỜI DÙNG</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew">
							<tr>
								<td width="89">&nbsp;</td>
								<td width="111"><b>Tên truy cập</b></td>
								<td width="364">
								<input type="text" name="txtUsername" size="30" class="input_text" value="<%=rsEdit("USERNAME")%>"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Mật khẩu</b></td>
								<td width="364">
								<input type="password" name="txtPassword" size="30" class="input_text" value="<%=rsEdit("PASSWORD")%>"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Nhập lại mật khẩu</b></td>
								<td width="364">
								<input type="password" name="txtPasswordAgain" size="30" class="input_text" value="<%=rsEdit("PASSWORD")%>"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Nhóm quyền</b></td>
								<td width="364">
								<select size="1" name="txtGroup" class="input_text">
								<%
									strSQL = "Select NAME, ID From USERGROUP"
									Call ListCombo(strSQL, rsEdit("GROUP_ID"))
								%>
								</select></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Họ và tên</b></td>
								<td width="364">
								<input type="text" name="txtFullname" size="30" class="input_text" value="<%=rsEdit("FULLNAME")%>"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="89"></td>
								<td height="14" width="111"><b>Ngày sinh</b></td>
								<td height="14" width="364">
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
								<td width="89">&nbsp;</td>
								<td width="475" colspan="2">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
									<tr>
										<td width="145">&nbsp;</td>
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