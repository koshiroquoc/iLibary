<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("notice")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>

<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../editor/fckeditor.asp" -->
<%
	txtCategory	= Request.Form("category")	
	If txtCategory = "schedule" Then				
		txtTitle	= Request.Form("txtTitle")
		If txtTitle = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtSummary	= Request.Form("txtSummary")
		If txtSummary = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtContent	= Request.Form("txtContent")
		If txtContent = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		strSQL = "SELECT * FROM SCHEDULE Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO SCHEDULE(ID,TITLE,SUMMARY,CONTENT,DATE_INFORM) VALUES("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtTitle,",") & CheckString(txtSummary,",")
		strSQL = strSQL & CheckString(txtContent,",")& CheckString(Now(),")")
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listschedule.asp")
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
					<td width="187" valign="top" background="../images/bg_menuleft.gif"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="4" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp; </b>
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font color="#FF0000" size="2">THÊM LỊCH TRỰC</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" name="frmSchedule" action="admin_addschedule.asp">
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b><font size="2">Tiêu đề lịch</font></b></td>
								<td width="459">
								<input type="text" name="txtTitle" size="27" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b><font size="2">Tóm tắt</font></b></td>
								<td width="459">
								<textarea rows="4" name="txtSummary" cols="51" class="input_text"></textarea></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91" valign="top"><b>
								<font size="2">Nội dung</font></b></td>
								<td height="14" width="459">
														
								<textarea rows="4" name="txtContent" cols="51" class="input_text"></textarea></td>
								<td height="14" width="9">
								</td>
							</tr>
								<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="14">&nbsp;</td>
								<td width="91">&nbsp;</td>
								<td width="459">
								<p align="center">
								<input type="submit" value="Tạo mới" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="schedule">
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