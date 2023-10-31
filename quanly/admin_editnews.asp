<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("news")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../editor/fckeditor.asp" -->
<%
	id	= Request.QueryString("id")	
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM NEWS WHERE ID="&id
	rsEdit.CursorType = 2
	rsEdit.LockType = 3
	rsEdit.Open strSQL, Conn
	
	txtCategory	= Request.Form("category")
	If txtCategory = "news" Then				
		txtTitle	= Request.Form("txtTitle")
		If txtTitle = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtSummary	= Request.Form("txtSummary")
		If txtSummary = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtCategory	= Request.Form("txtCategory")

		txtImage	= Request.Form("txtImage")
		If txtImage = "" Then
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
		
		txtHotNews	= Request.Form("txtHotNews")
				
		strSQL = "UPDATE NEWS SET "
		strSQL = strSQL & "TITLE ="&CheckString(txtTitle,",")
		strSQL = strSQL & "CATEGORY_ID="&CheckString(txtCategory,",")		
		strSQL = strSQL & "SUMMARY="&CheckString(txtSummary,",")		
		strSQL = strSQL & "CONTENT="&CheckString(txtContent,",")
		strSQL = strSQL & "AUTHOR="&CheckString(txtAuthor,",")
		strSQL = strSQL & "HOTNEWS="&CheckString(txtHotNews,",")		
		strSQL = strSQL & "IMAGE="&CheckString(txtImage,"")
		strSQL = strSQL & "WHERE ID="& id

		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listnews.asp")
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
								<td colspan="5" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp; <font color="#FF0000" size="2">HIỆU CHỈNH TIN TỨC</font></b></td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew">
							<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b>Tiêu đề</b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtTitle" size="35" class="input_text" value="<%=rsEdit("TITLE")%>"></td>
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
								<textarea rows="4" name="txtSummary" cols="54" class="input_text"><%=rsEdit("SUMMARY")%></textarea></td>
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
								<b>Thể loại tin</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtCategory" class="input_text">
								<%
									strSQL = "Select NAME, ID From CATEGORY_NEWS"
									Call ListCombo(strSQL, rsEdit("CATEGORY_ID"))
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
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b>Ảnh minh họa</b></td>
								<td width="97">
								<img border="0" width="89" height="104" name ="txtDisplay" src="<%=rsEdit("IMAGE")%>"></td>
								<td width="352">
								<b>
								<a href="JavaScript:openWindow2('admin_upload.asp?dir=news&win=pop&targetis=txtImage&show=txtDisplay')">Tải ảnh</a></b></td>
								<td width="9">
								&nbsp;</td>
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
									oFCKeditor.Value = rsEdit("CONTENT")
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
								<td width="14">
								&nbsp;</td>
								<td width="91">
								<b>Nguồn</b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtAuthor" size="30" class="input_text" value="<%=rsEdit("AUTHOR")%>"></td>
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
								<b>Tin nóng</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtHotNews" class="input_text">
								<%
									If rsEdit("HOTNEWS") = 0 Then
								%>
								<option value="1">Tin nóng</option>
								<option selected value="0">Không</option>
								<%
									Else
								%>
								<option value="0">Không</option>
								<option selected value="1">Tin nóng</option>
								<%
									End If
								%>
								</select></td>
								<td width="9">
								&nbsp;</td>
							</tr>
								<tr>
								<td width="14">&nbsp;</td>
								<td width="91">&nbsp;</td>
								<td width="459" colspan="2">
								<p align="center">
								<input type="submit" value="Cập nhật" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="news">
							<input type="hidden" name="txtImage" value="<%=rsEdit("IMAGE")%>">
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