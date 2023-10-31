<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("card")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	txtCategory	= Request.Form("category")	
	If txtCategory = "card" Then				
		txtClass	= Request.Form("txtClass")
		If txtClass = "All" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtClassName	= uCase(Request.Form("txtClassName"))
		If txtClassName = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		If Len(txtClassName) <> 2 Then
			If Len(txtClassName) <> 4 Then
				If Len(txtClassName) <> 6 Then	
					Response.Redirect("admin_error.asp?type=14")		
				End If
			End If		
		End If
		
		If Len(txtClassName) = 2 Then
			txtClassName = txtClass & txtClassName
		End If
		
		txtAccess = Request.Form("txtAccess")
		If txtAccess = "Excel" Then	
			txtFileName	= uCase(Request.Form("txtFileName"))
			If txtFileName = "" Then
				Response.Redirect("admin_error.asp?type=1")
			End If
		End If
		
		strSQL = "SELECT * FROM CARD WHERE CLASS_ID='" & txtClassName & "'"
		Set rsCheck = Server.CreateObject("ADODB.Recordset")
		rsCheck.CursorType = 2
		rsCheck.LockType = 3
		rsCheck.Open strSQL, Conn
		If Not rsCheck.Eof Then
			Response.Redirect("admin_error.asp?type=13")
		End If
		If txtAccess = "Excel" Then	
			Call ImportExcel(txtFileName,Conn)
		End If
		strSQL = "SELECT * FROM EXCEL_IMPORT"		
		Set rsExcel = Server.CreateObject("ADODB.Recordset")
		rsExcel.CursorType = 2
		rsExcel.LockType = 3
		rsExcel.Open strSQL, Conn
		Do While Not rsExcel.Eof
			strSQL = "SELECT * FROM CARD Order by ID Desc"		
			txtID = GetID(strSQL,Conn)
			
			txtCardID = ZenCardID(txtClassName)
			
			txtFirstName = rsExcel("FIRSTNAME")
			txtLastName = rsExcel("LASTNAME")
			txtBirthday = rsExcel("BIRTHDAY")		
			
			strSQL = "INSERT INTO CARD(ID,FIRSTNAME,LASTNAME,BIRTHDAY,CARD_ID,CLASS_ID,DATE_INFORM)Values("
			strSQL = strSQL & CheckString(txtID,",")& CheckString(txtFirstName,",")
			strSQL = strSQL & CheckString(txtLastName,",") & CheckString(txtBirthday,",")
			strSQL = strSQL & CheckString(txtCardID,",") & CheckString(txtClassName,",")
			strSQL = strSQL & CheckString(Now(),")")
			Conn.Execute strSQL
		rsExcel.MoveNext
		Loop
		
		strSQL = "DELETE * FROM EXCEL_IMPORT"
		Conn.Execute strSQL
		
		Conn.Close
		Set Conn = Nothing
		Response.Redirect("admin_card.asp")
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
								<b>&nbsp;<font size="2" color="#FF0000">TẠO THẺ THƯ VIỆN TỪ EXCEL</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" action="admin_importcard.asp" name="frmImportCard">
							<tr>
								<td width="72">&nbsp;</td>
								<td width="125"><b><font size="2">Loại độc giả</font></b></td>
								<td width="367">
								<font size="2">
								<select size="1" name="txtClass" class="input_text" onchange="JavaScript:cboChangeClass('txtClass');">
								<option selected value="All">-- Tất cả --</option>
								<%
									strSQL = "Select NAME,CATEGORY_ID From CATEGORY_CARD "
									Call ListCombo(strSQL,"All")
								%>
								</select></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="72">
								&nbsp;</td>
								<td width="125">
								<b><font size="2">Tên đơn vị /lớp</font></b></td>
								<td width="367">
								<font size="2">
								<input type="text" name="txtClassName" size="27" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="72">
								&nbsp;</td>
								<td width="125">
								<b><font size="2">Chọn loại</font></b></td>
								<td width="367">
								<font size="2">
								<select size="1" name="txtAccess" class="input_text">
								<option selected value="Access">Từ Access
								</option>
								<option value="Excel">Từ Excel</option>
								</select></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="72">&nbsp;</td>
								<td width="125"><b><font size="2">Chọn file import</font></b></td>
								<td width="367">
								<font size="2">
								<input type="file" name="txtFileName" size="27" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="573" colspan="4">
								&nbsp;</td>
								</tr>
								<tr>
								<td width="564" colspan="3">
								<p align="center">
								<input type="submit" value="Tạo mới" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="card">
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