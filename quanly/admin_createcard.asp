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
		
		
		txtCATACARDID = Request.Form("txtCATACARDID")	
		If txtCATACARDID = "All" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

			
		txtFirstName = Request.Form("txtFirstName")	
		If txtFirstName = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtLastName = Request.Form("txtLastName")	
		If txtLastName = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
			
		txtBirthday = Request.Form("txtDay") & "/" & Request.Form("txtMonth") & "/" & Request.Form("txtYear")
		
		txtClass = Request.Form("txtClass")	
		If txtClass = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtschoolyear= Request.Form("txtschoolyear")	
		If txtschoolyear = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If


		
		strSQL = "SELECT * FROM CARD Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
	
		txtCardID = ZenCardID(txtCATACARDID)

		strSQL = "INSERT INTO CARD(ID,CATACARD_ID,CARD_ID,FIRSTNAME,LASTNAME,BIRTHDAY,CLASS_ID,SCHOOLYEAR,DATE_INFORM)Values("
		strSQL = strSQL & CheckString(txtID,",")& CheckString(txtCATACARDID,",")
		strSQL = strSQL & CheckString(txtCardID,",") & CheckString(txtFirstName,",")
		strSQL = strSQL & CheckString(txtLastName,",") & CheckString(txtBirthday,",")
		strSQL = strSQL & CheckString(txtClass,",") & CheckString(txtschoolyear,",")
		strSQL = strSQL & CheckString(Now(),")")
		
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
					<td width="187" valign="top"><font size="2"><!--#INCLUDE FILE="admin_menu.asp" -->
					</font></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="4" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2">&nbsp; </font></b>
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>
								<font color="#FF0000" style="font-size: 11pt">TẠO THẺ THƯ VIỆN</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								</tr>
							<form method="POST" action="admin_createcard.asp">
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b><font size="2">Loại thẻ</font></b></td>
								<td width="320">
								<font size="2">
								<select size="1" name="txtCATACARDID" class="input_text">
									<option selected value="All">-- Tất cả --</option>
								<%
									strSQL = "Select DISTINCT CATEGORY_ID, NAME From CATEGORY_CARD"
									Call ListComboCARD(strSQL,"All")
								%>
								</select></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b><font size="2">Họ lót</font></b></td>
								<td width="320">
								<font size="2">
								<input type="text" name="txtFirstName" size="24" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b><font size="2">Tên</font></b></td>
								<td width="320">
								<font size="2">
								<input type="text" name="txtLastName" size="24" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b><font size="2">Ngày sinh</font></b></td>
								<td width="320">
								<font size="2">
								<select size="1" name="txtDay" class="input_text">
								<%
									Call ListNumber(01,31,"All")
								%>
								</select><select size="1" name="txtMonth" class="input_text">
								<%
									Call ListNumber(01,12,"All")
								%>
								</select><select size="1" name="txtYear" class="input_text">
								<%
									Call ListNumber(1945,2030,2004)
								%>
								</select></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b>
								<font size="2">Lớp</font></b></td>
								<td width="320">
								<font size="2">
								<input type="text" name="txtClass" size="24" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138" align="right">
								<b>
								<font size="2">Năm học</font></b></td>
								<td width="320">
								<font size="2">
								<input type="text" name="txtschoolyear" size="24" class="input_text"></font></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="106">
								&nbsp;</td>
								<td width="138">
								&nbsp;</td>
								<td width="320">
								&nbsp;</td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<font size="2">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
								</tr>
								<tr>
								<td width="573" colspan="4">
								&nbsp;</td>
								</tr>
								<tr>
								<td width="564" colspan="3">
								<p align="center">
								<font size="2">
								<input type="submit" value="Tạo mới" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></font></td>
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