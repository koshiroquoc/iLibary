<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("soft") = False Then
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
	If txtCategory = "softcategory" Then				
		txtName	= Request.Form("txtName")
		txtID	= Request.Form("txtID")
		If txtName = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
			
		If txtID <> "" Then
			strSQL = "UPDATE CATEGORY_SOFT SET "
			strSQL = strSQL & "NAME="& CheckString(txtName,"")
			strSQL = strSQL & "WHERE ID=" & cInt(txtID)
		Else
			strSQL = "SELECT * FROM CATEGORY_SOFT Order by ID Desc"		
			txtID = GetID(strSQL,Conn)			
			strSQL = "INSERT INTO CATEGORY_SOFT(ID,Name)Values("
			strSQL = strSQL & CheckString(txtID,",") & CheckString(txtName,")")
		End If	
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_softcategory.asp")
	Else
		If id <> "" Then
			Set rsEdit = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM CATEGORY_SOFT WHERE ID=" & cInt(id)
			rsEdit.CursorType = 2
			rsEdit.LockType = 3
			rsEdit.Open strSQL, Conn
			txtID = rsEdit("ID")
			txtName = rsEdit("NAME")
			rsEdit.Close
			Set rsEdit = Nothing
		End If		
		strSQL = "SELECT * FROM CATEGORY_SOFT ORDER BY ID DESC"		
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
					<td width ="797" height ="350" valign= "top">
					<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="6" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>
							<font size="2">&nbsp;</font><font color="#FF0000" size="2"> 
							THỂ LOẠI TIỆN ÍCH</font></b></td>
							</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="4">&nbsp;</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="4">&nbsp;</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<form method="POST" name="frmSoftCategory" action="admin_softcategory.asp">
								<td width="2%">&nbsp;</td>
								<td width="18%">&nbsp;</td>
								<td width="18%"><b><font size="2">Tên thể loại</font></b></td>
								<td width="23%">
								<font size="2">
								<input type="text" name="txtName" size="19" class="input_text" value="<%=txtName%>"></font></td>
								<td width="37%">
								<font size="2">
								<input type="submit" value="Cập nhật" name="B1" class="input_button"></font></td>
								<td width="2%">&nbsp;</td>
								<input type="hidden" name="category" value="softcategory">
								<input type="hidden" name="txtID" value="<%=txtID%>">
							</form>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="4">&nbsp;</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="4">
							<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
								<tr>
									<td width="103">&nbsp;</td>
									<td>
									<font size="2">
									<% ' If no data, don't display data table
										If Not rsSelect.Eof Then
									%>
									</font>
									<table border="1" width="100%" id="table5" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#999999" bordercolor="#000080">
										<tr>
											<td width="32" align="center" bgcolor="#EEEEEE" height="20"><b>
											<font size="2">STT</font></b></td>
											<td align="center" bgcolor="#EEEEEE" height="20"><b>
											<font size="2">Tên thể loại</font></b></td>
											<td width="91" align="center" bgcolor="#EEEEEE" height="20">&nbsp;</td>
										</tr>
										<%
											Dim iCount
											iCount = 1
											Do While Not rsSelect.Eof 
										%>
										<tr>
											<td width="32" align="center">
											<font size="2"><%=iCount%></font></td>
											<td><font size="2">&nbsp;<%=rsSelect("NAME")%></font></td>
											<td width="91">
											<p align="center">
											<a href="admin_softcategory.asp?id=<%=rsSelect("ID")%>">
											<font size="2">Sửa</font></a><font size="2">&nbsp; |&nbsp;</font><a href="admin_delete.asp?category=softcategory&id=<%=rsSelect("ID")%>"><font size="2"> Xóa</font></a></td>
										</tr>
										<%
											iCount = iCount + 1
											rsSelect.MoveNext											
											Loop
											rsSelect.Close
											Set rsSelect = Nothing
										%>
									</table>
									<font size="2">
									<%
										End If
									%>
									</font>
									</td>
									<td width="57">&nbsp;</td>
								</tr>
							</table>
							</td>
							<td width="2%">&nbsp;</td>
						</tr>
						<tr>
							<td width="2%">&nbsp;</td>
							<td colspan="4">&nbsp;</td>
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