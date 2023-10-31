<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("card") = False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If
	End If	
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	id	= Request.QueryString("id")
	txtCategory	= Request.Form("category")	
	If txtCategory = "ologycategory" Then				
		txtID	= Request.Form("txtID")

		txtCategoryID = Request.Form("txtCategoryID")
		If txtCategoryID = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtName	= Request.Form("txtName")
		If txtName = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
			
		If txtID <> "" Then
			strSQL = "UPDATE OLOGY SET "
			strSQL = strSQL & "OLOGY_ID="& CheckString(txtCategoryID,",")
			strSQL = strSQL & "NAME="& CheckString(txtName,"")
			strSQL = strSQL & "WHERE ID="& cInt(txtID)
		Else
			Set rsCheck = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM OLOGY WHERE OLOGY_ID='" & Trim(txtCategoryID) & "'"
			rsCheck.Open strSQL, Conn,3,1
			If Not rsCheck.Eof Then			
				rsCheck.Close
				Set rsCheck = Nothing
				Response.Redirect("admin_error.asp?type=2")
			Else
				rsCheck.Close
				Set rsCheck = Nothing

				strSQL = "SELECT * FROM OLOGY Order by ID Desc"		
				txtID = GetID(strSQL,Conn)
						
				strSQL = "INSERT INTO OLOGY(ID,OLOGY_ID,NAME)Values("
				strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCategoryID,",")
				strSQL = strSQL & CheckString(txtName,")")
			End If	
		End If	
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_ologycategory.asp")
	Else
		If id <> "" Then
			Set rsEdit = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM OLOGY WHERE ID=" & cInt(id)
			rsEdit.CursorType = 2
			rsEdit.LockType = 3
			rsEdit.Open strSQL, Conn
			txtID = rsEdit("ID")
			txtName = rsEdit("NAME")
			txtCategoryID = rsEdit("OLOGY_ID")
			rsEdit.Close
			Set rsEdit = Nothing
		End If		
		strSQL = "SELECT * FROM OLOGY ORDER BY ID DESC"		
		Set rsSelect = Server.CreateObject("ADODB.Recordset")
		rsSelect.CursorType = 2
		rsSelect.LockType = 3
		rsSelect.Open strSQL, Conn
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

<table border="1" width="760" id="table1" bordercolordark="#808080" cellspacing="0" cellpadding="0" bordercolorlight="#D5F1FF">
	<tr>
		<td>
		<div align="center">
			<table border="0" width="760" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td colspan="2"><!--#INCLUDE FILE="admin_header.asp" --></td>
				</tr>
				<tr>
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="573" height="350" valign= "top">
					<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
						<tr>
								<td colspan="5" background="../images/bg_title.gif" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px">
								<b>&nbsp; THÔNG TIN NGÀNH HỌC</b></td>
							</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="5">&nbsp;</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%" height="13"></td>
							<td colspan="5" height="13"></td>
							<td width="2%" height="13"></td>
						</tr>
						<form method="POST" name="frmBookCategory" action="admin_ologycategory.asp">
						<tr>
								<td width="2%">&nbsp;</td>
								<td width="16%">&nbsp;</td>
								<td width="13%"><b>Mã&nbsp; ngành</b></td>
								<td width="5%">
								<!--webbot bot="Validation" s-data-type="String" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="2" --><input type="text" name="txtCategoryID" size="2" class="input_text" value="<%=txtCategoryID%>" maxlength="2" style="text-align: center"></td>
								<td width="22%">
								&nbsp;(2 ký tự đại diện)</td>
								<td width="41%">
								&nbsp;</td>
								<td width="2%">&nbsp;</td>
								</tr>
						<tr>
								<td width="100%" colspan="7" height="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								<tr>
								<td width="2%">&nbsp;</td>
								<td width="16%">&nbsp;</td>
								<td width="13%"><b>Tên ngành</b></td>
								<td width="27%" colspan="2">
								<input type="text" name="txtName" size="19" class="input_text" value="<%=txtName%>"></td>
								<td width="41%">
								<input type="submit" value="Cập nhật" name="B1" class="input_button"></td>
								<td width="2%">&nbsp;</td>
								<input type="hidden" name="category" value="ologycategory">
								<input type="hidden" name="txtID" value="<%=txtID%>">
							</form>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="5">&nbsp;</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="5">
							<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
								<tr>
									<td width="88">&nbsp;</td>
									<td>
									<% ' If no data, don't display data table
										If Not rsSelect.Eof Then
									%>
									<table border="1" width="100%" id="table5" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#999999" bordercolor="#000080">
										<tr>
											<td width="32" align="center" bgcolor="#EEEEEE" height="20"><b>STT</b></td>
											<td align="center" bgcolor="#EEEEEE" height="20" width="64">
											<b>Mã ngành 
											</b></td>
											<td align="center" bgcolor="#EEEEEE" height="20"><b>Tên 
											ngành</b></td>
											<td width="62" align="center" bgcolor="#EEEEEE" height="20">&nbsp;</td>
										</tr>
										<%
											Dim iCount
											iCount = 1
											Do While Not rsSelect.Eof 
										%>
										<tr>
											<td width="32" align="center"><%=iCount%></td>
											<td width="64" align="center"><%=rsSelect("OLOGY_ID")%></td>
											<td>&nbsp;<%=rsSelect("NAME")%></td>
											<td width="62">
											<p align="center">
											<a href="admin_ologycategory.asp?id=<%=rsSelect("ID")%>">Sửa</a>&nbsp; |&nbsp;<a href="admin_delete.asp?category=ologycategory&id=<%=rsSelect("ID")%>"> Xóa</a></td>
										</tr>
										<%
											iCount = iCount + 1
											rsSelect.MoveNext											
											Loop
											rsSelect.Close
											Set rsSelect = Nothing
										%>
									</table>
									<%
										End If
									%>
									</td>
									<td width="129">&nbsp;</td>
								</tr>
							</table>
							</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="5">&nbsp;</td>
							<td width="2%">&nbsp;</td>
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

</div>

</body>
</html>
<% End If %>