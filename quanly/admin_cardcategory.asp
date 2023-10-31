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
	If txtCategory = "cardcategory" Then				
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
			strSQL = "UPDATE CATEGORY_CARD SET "
			strSQL = strSQL & "CATEGORY_ID="& CheckString(txtCategoryID,",")
			strSQL = strSQL & "NAME="& CheckString(txtName,"")
			strSQL = strSQL & "WHERE ID="& cInt(txtID)
		Else
			Set rsCheck = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM CATEGORY_CARD WHERE CATEGORY_ID='" & Trim(txtCategoryID) & "'"
			rsCheck.Open strSQL, Conn,3,1
			If Not rsCheck.Eof Then			
				rsCheck.Close
				Set rsCheck = Nothing
				Response.Redirect("admin_error.asp?type=2")
			Else
				rsCheck.Close
				Set rsCheck = Nothing

				strSQL = "SELECT * FROM CATEGORY_CARD Order by ID Desc"		
				txtID = GetID(strSQL,Conn)
						
				strSQL = "INSERT INTO CATEGORY_CARD(ID,CATEGORY_ID,NAME)Values("
				strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCategoryID,",")
				strSQL = strSQL & CheckString(txtName,")")
			End If	
		End If	
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_cardcategory.asp")
	Else
		If id <> "" Then
			Set rsEdit = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM CATEGORY_CARD WHERE ID=" & cInt(id)
			rsEdit.CursorType = 2
			rsEdit.LockType = 3
			rsEdit.Open strSQL, Conn
			txtID = rsEdit("ID")
			txtName = rsEdit("NAME")
			txtCategoryID = rsEdit("CATEGORY_ID")
			rsEdit.Close
			Set rsEdit = Nothing
		End If		
		strSQL = "SELECT * FROM CATEGORY_CARD ORDER BY ID DESC"		
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
					<td width ="797" height="350" valign= "top">
					<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="7" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>
							<font size="2">&nbsp;&nbsp;</font></b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<b><font size="2">&nbsp;</font><font color="#FF0000" size="2">LOẠI ĐỘC GIẢ</font></b></td>
							</tr>
						<tr>
							<td colspan="7">&nbsp;</td>
						</tr>
						<tr>
							<td width="1%" height="13"></td>
							<td colspan="5" height="13"></td>
							<td width="1%" height="13"></td>
						</tr>
						<form method="POST" name="frmAddNew" action="admin_cardcategory.asp">
						<tr>
								<td width="1%">&nbsp;</td>
								<td width="18%">&nbsp;</td>
								<td width="13%"><b><font size="2">Mã&nbsp; loại</font></b></td>
								<td width="7%">
								<font size="2">
								<!--webbot bot="Validation" s-data-type="String" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="2" --><input type="text" name="txtCategoryID" size="2" class="input_text" value="<%=txtCategoryID%>" maxlength="2" style="text-align: center"></font></td>
								<td width="53%" colspan="2">
								<font size="2">&nbsp;2 ký tự đại diện, ví dụ: NV, GV, QL, HS</font></td>
								<td width="1%">&nbsp;</td>
								</tr>
						<tr>
								<td width="100%" colspan="7" height="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								<tr>
								<td width="1%">&nbsp;</td>
								<td width="18%">&nbsp;</td>
								<td width="13%"><b><font size="2">Tên</font></b></td>
								<td width="28%" colspan="2">
								<font size="2">
								<input type="text" name="txtName" size="19" class="input_text" value="<%=txtName%>"></font></td>
								<td width="36%">
								<font size="2">
								<input type="submit" value="Cập nhật" name="B1" class="input_button"></font></td>
								<td width="1%">&nbsp;</td>
								<input type="hidden" name="category" value="cardcategory">
								<input type="hidden" name="txtID" value="<%=txtID%>">
							</form>
						</tr>
						<tr>
							<td width="1%">&nbsp;</td>
							<td colspan="5">&nbsp;</td>
							<td width="1%">&nbsp;</td>
						</tr>
						<tr>
							<td width="1%">&nbsp;</td>
							<td colspan="5">
							<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
								<tr>
									<td width="107">&nbsp;</td>
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
											<td align="center" bgcolor="#EEEEEE" height="20" width="72">
											<b><font size="2">Mã loại</font></b></td>
											<td align="center" bgcolor="#EEEEEE" height="20">
											<b><font size="2">Tên</font></b></td>
											<td width="68" align="center" bgcolor="#EEEEEE" height="20">&nbsp;</td>
										</tr>
										<%
											Dim iCount
											iCount = 1
											Do While Not rsSelect.Eof 
										%>
										<tr>
											<td width="32" align="center">
											<font size="2"><%=iCount%></font></td>
											<td width="72" align="center">
											<font size="2"><%=rsSelect("CATEGORY_ID")%></font></td>
											<td><font size="2">&nbsp;<%=rsSelect("NAME")%></font></td>
											<td width="68">
											<p align="center">
											<a href="admin_cardcategory.asp?id=<%=rsSelect("ID")%>">
											<font size="2">Sửa</font></a><font size="2">&nbsp; |&nbsp;</font><a href="admin_delete.asp?category=cardcategory&id=<%=rsSelect("ID")%>"><font size="2"> Xóa</font></a></td>
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
									<td width="140">&nbsp;</td>
								</tr>
							</table>
							</td>
							<td width="1%">&nbsp;</td>
						</tr>
						<tr>
							<td width="1%">&nbsp;</td>
							<td colspan="5">&nbsp;</td>
							<td width="1%">&nbsp;</td>
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