<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("book")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>
<%
	id	= Request.QueryString("id")	
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM card WHERE ID="&id
	rsEdit.CursorType = 2
	rsEdit.LockType = 3
	rsEdit.Open strSQL, Conn
	
	txtCategoryOld = rsEdit("BOOK_ID")
	
	txtCategory	= Request.Form("category")
	If txtCategory = "book" Then				
		txtTitle	= Request.Form("txtTitle")

		txtSummary	= Request.Form("txtSummary")
		
		txtGenre	= Request.Form("txtGenre")
		
		txtCategoryNew	= Request.Form("txtCategory")

		txtLanguage	= Request.Form("txtLanguage")

		txtImage	= Request.Form("txtImage")

		txtPublisher	= Request.Form("txtPublisher")
		txtYearPublish	= Request.Form("txtYearPublish")
		txtVolume	= Request.Form("txtVolume")
		txtAmount	= Request.Form("txtAmount")
		
		txtAuthor	= Request.Form("txtAuthor")
						
		If Left(txtCategoryOld,3) <> txtCategoryNew Then
			txtBookID = ZenBookID(txtCategoryNew)
		Else
			txtBookID = txtCategoryOld
		End If	
		
		strSQL = "UPDATE BOOK SET "
		strSQL = strSQL & "BOOK_ID="&CheckString(txtBookID,",")
		strSQL = strSQL & "NAME="&CheckString(txtTitle,",")		
		strSQL = strSQL & "AUTHOR="&CheckString(txtAuthor,",")				
		strSQL = strSQL & "SUMMARY="&CheckString(txtSummary,",")		
		strSQL = strSQL & "GENRE="&CheckString(txtGenre,",")				
		strSQL = strSQL & "VOLUME="&CheckString(txtVolume,",")
		strSQL = strSQL & "AMOUNT="&CheckString(txtAmount,",")		
		strSQL = strSQL & "PUBLISHER="&CheckString(txtPublisher,",")
		strSQL = strSQL & "YEAR_PUBLISH="&CheckString(txtYearPublish,",")		
		strSQL = strSQL & "LANGUAGE="&CheckString(txtLanguage,",")		
		strSQL = strSQL & "IMAGE="&CheckString(txtImage,"")
		strSQL = strSQL & "WHERE ID="& id

		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listcard.asp")
	Else
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
<script language="javascript">
	function CheckInput(){		
		if(document.frmAddNew.txtTitle.value == ""){
			alert("Bạn chưa tên sách!");
			document.frmAddNew.txtTitle.focus();
			return;
			}
		if(document.frmAddNew.txtTitle.value.length < 6){
			alert("Tên sách quá ngắn, ít nhất phải 6 ký tự!");
			document.frmAddNew.txtTitle.focus();
			return;
			}
		if(document.frmAddNew.txtSummary.value == ""){
			alert("Bạn chưa nhập giới thiệu về sách!");
			document.frmAddNew.txtSummary.focus();
			return;
		}
		if(document.frmAddNew.txtSummary.value.length < 10){
			alert("Tóm tắt sách quá ngắn, ít nhất phải 10 ký tự!");
			document.frmAddNew.txtSummary.focus();
			return;
		}		
		if(document.frmAddNew.txtAuthor.value == ""){
			alert("Bạn chưa nhập tên tác giả!");
			document.frmAddNew.txtContent.focus();
			return;
		}
		if(document.frmAddNew.txtAuthor.value.length < 6){
			alert("Tên tác giả quá ngắn, ít nhất phải 6 ký tự!");
			document.frmAddNew.txtContent.focus();
			return;
		}
		document.frmAddNew.submit();
	}
</script>
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
								<b>&nbsp;</b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;<font color="#FF0000" size="2">HIỆU CHỈNH SÁCH</font></b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew">
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b>Tên sách</b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtTitle" size="30" class="input_text" value="<%=rsEdit("NAME")%>"></td>
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
								<b>Ảnh minh họa</b></td>
								<td width="97">
								<img border="0" width="89" height="104" name ="txtDisplay" src="<%=rsEdit("IMAGE")%>"></td>
								<td width="362">
								<b>
								<a href="JavaScript:openWindow2('admin_upload.asp?dir=book&win=pop&targetis=txtImage&show=txtDisplay')">
								Tải ảnh</a></b></td>
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
								<b>Thể loại</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtGenre" class="input_text">
								<%
									strSQL = "Select NAME, ID From CATEGORY_GENRE"
									Call ListCombo(strSQL, rsEdit("GENRE"))
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
								<b>Lĩnh vực</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtCategory" class="input_text">
								<%
									strSQL = "Select NAME, CATEGORY_ID From CATEGORY_BOOK"
									Call ListCombo(strSQL, Left(rsEdit("BOOK_ID"),3))
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
								<b>Số lượng</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtAmount" class="input_text">
								<%
									Call ListNumber(1,20, cInt(rsEdit("AMOUNT")))
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
								<b>Tác giả</b></td>
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
								<td height="14" width="14"></td>
								<td height="14" width="91"><b>Ngôn 
								ngữ</b></td>
								<td height="14" width="459" colspan="2">
								<select size="1" name="txtLanguage" class="input_text">
								<%
									strSQL = "Select NAME, ID From LANGUAGE"
									Call ListCombo(strSQL, rsEdit("LANGUAGE"))
								%>
								</select></td>
								<td height="14" width="9">
								</td>
							</tr>
								<tr>
								<td width="573" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91"><b>Nhà 
								xuất bản</b></td>
								<td height="14" width="459" colspan="2">
								<select size="1" name="txtPublisher" class="input_text">
								<%
									strSQL = "Select NAME, ID From PUBLISHER"
									Call ListCombo(strSQL, rsEdit("PUBLISHER"))
								%>
								</select></td>
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
								<b>Năm xuất bản</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtYearPublish" class="input_text">
								<%
									Call ListNumber(1945,2006,rsEdit("YEAR_PUBLISH"))
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
								<b>Lần xuất bản</b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtVolume" class="input_text">
								<%
									Call ListNumber(1,5, rsEdit("VOLUME"))
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
								<td width="14">&nbsp;</td>
								<td width="91">&nbsp;</td>
								<td width="459" colspan="2">
								<p align="center">
								<input type="button" value="Cập nhật" name="B2" class="input_button" onclick="JavaScript:CheckInput();">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="book">
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