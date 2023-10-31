<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
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
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	txtCategory	= Request.Form("category")	
	If txtCategory = "book" Then				
		txtTitle	= Trim(Request.Form("txtTitle"))

		txtSummary	= Trim(Request.Form("txtSummary"))
		
		txtGenre	= Request.Form("txtGenre")

		txtCategory	= Request.Form("txtCategory")

		txtLanguage	= Request.Form("txtLanguage")

		txtImage	= Request.Form("txtImage")
		If txtImage = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtPublisher	= Request.Form("txtPublisher")
		txtYearPublish	= Request.Form("txtYearPublish")
		txtVolume	= Request.Form("txtVolume")
		txtAmount	= Request.Form("txtAmount")
		
		txtAuthor	= Trim(Request.Form("txtAuthor"))
				
		strSQL = "SELECT * FROM BOOK Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		txtBookID = ZenBookID(txtCategory)
		
		strSQL = "INSERT INTO BOOK(ID,BOOK_ID,NAME,AUTHOR,SUMMARY,GENRE,VOLUME,AMOUNT,PUBLISHER,YEAR_PUBLISH,LANGUAGE,IMAGE,DATE_INFORM) VALUES("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtBookID,",") & CheckString(txtTitle,",")
		strSQL = strSQL & CheckString(txtAuthor,",") & CheckString(txtSummary,",") & CheckString(txtGenre,",") & CheckString(txtVolume,",")
		strSQL = strSQL & CheckString(txtAmount,",") & CheckString(txtPublisher,",") & CheckString(txtYearPublish,",")
		strSQL = strSQL & CheckString(txtLanguage,",") & CheckString(txtImage,",") &  CheckString(Now(),")")
		Conn.Execute strSQL
		Conn.Close
		Set Conn = Nothing

		Response.Redirect("admin_listbook.asp")
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
						<table border="0" width="661" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="5" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b>&nbsp;</b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font color="#FF0000">&nbsp;</font><font color="#FF0000" size="3">THÊM SÁCH MỚI</font></b></td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew" action="admin_addbook.asp">
							<tr>
								<td width="14">&nbsp;</td>
								<td width="91"><b><font size="2">Tên sách</font></b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtTitle" size="30" class="input_text"></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Tóm tắt</font></b></td>
								<td width="459" colspan="2">
								<textarea rows="4" name="txtSummary" cols="54" class="input_text"></textarea></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Ảnh minh họa</font></b></td>
								<td width="97">
								<img border="0" width="86" height="95" name ="txtDisplay" src=""></td>
								<td width="362">
								<b>
								<a href="JavaScript:openWindow2('admin_upload.asp?dir=book&win=pop&targetis=txtImage&show=txtDisplay')">
								Tải ảnh</a></b></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Thể loại</font></b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtGenre" class="input_text">
								<%
									strSQL = "Select NAME, ID From CATEGORY_GENRE"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Lĩnh vực</font></b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtCategory" class="input_text">
								<%
									strSQL = "Select NAME, CATEGORY_ID From CATEGORY_BOOK"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Số lượng</font></b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtAmount" class="input_text">
								<%
									Call ListNumber(1,20, "All")
								%>
								</select></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Tác giả</font></b></td>
								<td width="459" colspan="2">
								<input type="text" name="txtAuthor" size="30" class="input_text"></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91"><b><font size="2">Ngôn 
								ngữ</font></b></td>
								<td height="14" width="459" colspan="2">
								<select size="1" name="txtLanguage" class="input_text">
								<%
									strSQL = "Select NAME, ID From LANGUAGE"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td height="14" width="97">
								</td>
							</tr>
								<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="14"></td>
								<td height="14" width="91"><b><font size="2">Nhà 
								xuất bản</font></b></td>
								<td height="14" width="459" colspan="2">
								<select size="1" name="txtPublisher" class="input_text">
								<%
									strSQL = "Select NAME, ID From PUBLISHER"
									Call ListCombo(strSQL, "All")
								%>
								</select></td>
								<td height="14" width="97">
								</td>
							</tr>
								<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Năm xuất bản</font></b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtYearPublish" class="input_text">
								<%
									Call ListNumber(1975,2006,"All")
								%>
								</select></td>
								<td width="97">&nbsp;
								</td>
							</tr>
								<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td width="14">&nbsp;
								</td>
								<td width="91">
								<b><font size="2">Lần xuất bản</font></b></td>
								<td width="459" colspan="2">
								<select size="1" name="txtVolume" class="input_text">
								<%
									Call ListNumber(1,5,"All")
								%>
								</select></td>
								<td width="97">&nbsp;
								</td>
							</tr>
							<tr>
								<td width="661" colspan="5">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="14">&nbsp;</td>
								<td width="91">&nbsp;</td>
								<td width="459" colspan="2">
								<p align="center">
								<font size="2">
								<input type="button" value="Tạo mới" name="B2" class="input_button" onclick="JavaScript:CheckInput();"></font>&nbsp;
								<font size="2">
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></font></td>
								<td width="97">&nbsp;
								</td>
								</tr>
								<input type="hidden" name="category" value="book">
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