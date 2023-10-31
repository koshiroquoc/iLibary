<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("doc")= False Then
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
	If txtCategory = "document" Then				
		txtTitle	= Request.Form("txtTitle")
		If txtTitle = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtSummary	= Request.Form("txtSummary")
		If txtSummary = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtCategory	= Request.Form("txtCategory")

		txtAuthor	= Request.Form("txtAuthor")
		If txtAuthor = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
		
		txtContent	= Request.Form("txtContent")
		If txtContent = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If
				
		strSQL = "SELECT * FROM DOCUMENT Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		txtDocumentID = ZenDocID(txtCategory)
		
		strSQL = "INSERT INTO DOCUMENT(ID,DOCUMENT_ID,TITLE,CATEGORY_ID,SUMMARY,CONTENT,AUTHOR,FILE_PATH,DATE_INFORM) VALUES("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtDocumentID,",") & CheckString(txtTitle,",")
		strSQL = strSQL & CheckString(txtCategory,",") & CheckString(txtSummary,",") & CheckString(txtContent,",")
		strSQL = strSQL & CheckString(txtAuthor,",") & CheckString(txtFilePath,",") &  CheckString(Now(),")")
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listdoc.asp")
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
								<td colspan="4" height="31">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;<font color="#FF0000" size="2"> </font></b>
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font color="#FF0000" size="2"> THÊM TÀI LIỆU</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew" action="admin_adddoc.asp">
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Tên tài liệu</b></td>
								<td width="459">
								<input type="text" name="txtTitle" size="30" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Tóm tắt</b></td>
								<td width="459">
								<textarea rows="4" name="txtSummary" cols="54" class="input_text"></textarea></td>
								<td width="9">&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Thể loại</b></td>
								<td width="459">
								<select size="1" name="txtCategory" class="input_text">
								<%
									strSQL = "Select NAME, CATEGORY_ID From CATEGORY_DOCUMENT"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td width="9">&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b>Tác giả</b></td>
								<td width="459">
								<input type="text" name="txtAuthor" size="30" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Nội dung</b></td>
								<td width="459">
								<textarea rows="9" name="txtContent" cols="56" class="input_text"></textarea>								
								</td>
								<td width="9">&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>File đính kèm</b></td>
								<td width="459">
								<input type="text" name="txtFilePath" size="30" class="input_text">
								<a href="JavaScript:openWindow2('admin_uploadfile.asp?dir=doc&win=pop&targetis=txtFilePath')"><b>Tải file</b></a></td>
								<td width="9">&nbsp;</td>
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
								<td width="9">&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="document">							
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