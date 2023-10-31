<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("relax")= False Then
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
	If txtCategory = "relax" Then				
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

		txtAuthor	= Request.Form("txtAuthor")
		If txtAuthor = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtImage	= Request.Form("txtImage")
		txtCategory	= Request.Form("txtCategory")
		
		strSQL = "SELECT * FROM RELAX Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO RELAX(ID,TITLE,CATEGORY_ID,SUMMARY,CONTENT,AUTHOR,IMAGE,DATE_INFORM) VALUES("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtTitle,",")
		strSQL = strSQL & CheckString(txtCategory,",") & CheckString(txtSummary,",")
		strSQL = strSQL & CheckString(txtContent,",")& CheckString(txtAuthor,",")
		strSQL = strSQL & CheckString(txtImage,",") & CheckString(Now(),")")
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listrelax.asp")
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
								<td colspan="5" height="37">
								<p style="margin-top: 2px; margin-bottom: 2px">
								<font color="#FF0000">
								<b>&nbsp; </b></font>
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<font color="#FF0000">
								<b><font size="2">THÊM GIẢI TRÍ</font></b></font></td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew" action="admin_addrelax.asp">
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Tiêu đề</b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtTitle" size="27" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b>Tóm tắt</b></td>
								<td width="459" colspan="2">
								<textarea rows="4" name="txtSummary" cols="51" class="input_text"></textarea></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Thể loại</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtCategory" class="input_text">
								<%
									strSQL = "Select NAME, ID From CATEGORY_RELAX"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91"><b>Ảnh minh họa</b></td>
								<td height="14" width="105">
								<img border="0" width="96" height="102" name="txtDisplay"></td>
								<td height="14" width="354">
								<b>
								<a href="JavaScript:openWindow2('admin_upload.asp?dir=relax&win=pop&targetis=txtImage&show=txtDisplay')">Tải ảnh</a></b></td>
								<td height="14" width="9">
								</td>
							</tr>
								<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91" valign="top"><b>Nội dung</b></td>
								<td height="14" width="459" colspan="2">
								<%
									Set oFCKeditor = New FCKeditor
									oFCKeditor.BasePath	= "/editor/"
									oFCKeditor.Height = 300
									oFCKeditor.Create "txtContent"
								%>								</td>
								<td height="14" width="9">
								</td>
							</tr>
								<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Nguồn</b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtAuthor" size="27" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="14">&nbsp;</td>
								<td width="91">&nbsp;</td>
								<td width="459" colspan="2">
								<p align="center">
								<input type="submit" value="Tạo mới" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="relax">
							<input type="hidden" name="txtImage" value=" ">
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