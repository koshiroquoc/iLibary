<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
'	If Session("Username")= "" Then
'		Response.Redirect("admin_login.asp")
'	End If	
'	If Session("Admin") = False Then
'		Response.Redirect("admin_error.asp?type=5")
'	End If	
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	txtCategory	= Request.Form("category")	
	If txtCategory = "user" Then				
		
		txtUser	= Request.Form("txtUser")
		If txtUser = "All" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		Set rsCheck = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM MODULE WHERE USERNAME ='" & txtUser & "'"
		rsCheck.Open strSQL, Conn,3,1
		If Not rsCheck.Eof Then
			strSQL = "DELETE * FROM MODULE WHERE USERNAME ='" & txtUser & "'"
			Conn.Execute strSQL
		End If	
		rsCheck.Close
		Set rsCheck = Nothing
		
		For i = 1 To Request.Form("Mid").Count
			txtFunction = Request.Form("Mid")(i)
			strSQL = "SELECT * FROM MODULE Order by ID Desc"		
			txtID = GetID(strSQL,Conn)
				
			strSQL = "INSERT INTO MODULE (ID,FUNCTION_ID,USERNAME)Values("
			strSQL = strSQL & CheckString(txtID,",") & CheckString(txtFunction,",")
			strSQL = strSQL & CheckString(txtUser,")")
			Conn.Execute strSQL
		Next
			
		Conn.Close
		Set Conn = Nothing
		Response.Redirect("admin_done.asp?page=admin_user.asp")		
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
							<td colspan="6" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>&nbsp; 
							</b>
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<b><font color="#FF0000" size="2">PHÂN QUYỀN CHO NGƯỜI DÙNG</font></b></td>
							</tr>
							<tr>
								<td colspan="6" height="13"></td>
							</tr>
							<form method="POST" name="frmList" action="admin_setuser.asp">
							<tr>
								<td colspan="6" height="13"></td>
							</tr>
							<tr>
								<td width="16">&nbsp;</td>
								<td width="82">&nbsp;</td>
								<td width="121"><p align="center"><b>
								<font size="2">Chọn người 
								dùng</font></b></td>
								<td width="32">
								<font size="2">
								<select size="1" name="txtUser" class="input_text" onchange="JavaScript:cboChange('txtUser');">
								<option selected value="All">-- Tất cả --</option>
								<%
									strSQL = "Select FULLNAME, USERNAME From USER WHERE LEVEL<>1 ORDER BY FULLNAME"
									Call ListCombo(strSQL, Request.Form("txtCategoryFilter"))
								%>
								</select></font></td>
								<td width="293">&nbsp;</td>
								<td width="11">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="6">
								&nbsp;</td>
							</tr>
							<tr>
								<td colspan="6">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
									<%
										Set rsList = Server.CreateObject("ADODB.Recordset")
										strSQL = "SELECT * FROM FUNCTION Order by NAME ASC"
										rsList.Open strSQL, Conn,3,1
										txtCategoryFilter = Request.Form("txtCategoryFilter")												
										Do While Not rsList.Eof
									%>
									<tr>
										<td width="209">&nbsp;</td>
										<td width="27">
										<p align="center">
										<font size="2">
										<%
											strSQL = "SELECT * FROM MODULE WHERE FUNCTION_ID=" & rsList("ID")
											strSQL = strSQL & " AND USERNAME='"& txtCategoryFilter & "'"
											Set rsShow = Server.CreateObject("ADODB.Recordset")
											rsShow.Open strSQL, Conn,3,1
											If Not rsShow.Eof Then
										%>
										</font>
										<input type="checkbox" name="Mid" value="<%=rsList("ID")%>" checked></td>
										<%
											Else
										%>
										<input type="checkbox" name="Mid" value="<%=rsList("ID")%>"><font size="2"></td>
										<%
											End If
										%>
										</font>
										<td><font size="2">&nbsp;<%=rsList("NAME")%></font></td>
									</tr>
									<%
										rsList.MoveNext
										Loop
										Conn.Close
										Set Conn = Nothing
									%>												
								</table>
								</td>
							</tr>
							<tr>
								<td colspan="6">
								&nbsp;</td>
							</tr>
							<tr>
								<td colspan="6">
								<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
									<tr>
										<td width="169">&nbsp;</td>
										<td width="174">
										<p align="center">
													<font size="2">
													<input type="submit" value="Cập nhật" name="B3" class="input_button">
										<input type="reset" value=" Hủy bỏ " name="B4" class="input_button"></font></td>
										<td>&nbsp;</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td colspan="6">
								&nbsp;</td>
							</tr>
								<input type="hidden" name="category" value="user">
							</form>
						</table>
					</div>
					</td>
				</tr>
				<form method="POST" name="frmFilter" action="admin_setuser.asp">
					<input type="hidden" name="txtCategoryFilter" value="">
				</form>						
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