<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("vote")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>
<%
	txtCategory	= Request.Form("category")	
	If txtCategory = "vote" Then				
		txtQuestion	= Request.Form("txtQuestion")
		If txtQuestion = "" Then
			Response.Redirect("admin_error.asp?type=1")
		End If

		txtStatus	= Request.Form("txtStatus")
		If txtStatus = 1 Then
			strSQL = "SELECT * FROM QUESTION"
			Call UpdateField(strSQL)
		End If

		strSQL = "SELECT * FROM QUESTION Order by ID Desc"		
		txtQuestionID = GetID(strSQL,Conn)

		strSQL = "INSERT INTO QUESTION(ID,NAME,STATUS,DATE_INFORM) VALUES("
		strSQL = strSQL & CheckString(txtQuestionID,",") & CheckString(txtQuestion,",")
		strSQL = strSQL & CheckString(txtStatus,",")
		strSQL = strSQL & CheckString(Now(),")")
		Conn.Execute strSQL

		For i = 1 To Request.Form("txtAnswer").Count
			txtTitle = Request.Form("txtAnswer")(i)
			If txtTitle <> "" Then
				strSQL = "SELECT * FROM VOTE Order by ID Desc"		
				txtVoteID = GetID(strSQL,Conn)
				
				strSQL = "INSERT INTO VOTE(ID,CATEGORY_ID,NAME,DATE_INFORM) VALUES("
				strSQL = strSQL & CheckString(txtVoteID,",") & CheckString(txtQuestionID,",")
				strSQL = strSQL & CheckString(txtTitle,",")
				strSQL = strSQL & CheckString(Now(),")")
				Conn.Execute strSQL
			End If
			Next
		Conn.Close
		Set Conn = Nothing
		Response.Redirect("admin_listvote.asp")
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

<table border="1" width="760" id="table1" bordercolordark="#808080" cellspacing="0" cellpadding="0" bordercolorlight="#D5F1FF">
	<tr>
		<td>
		<div align="center">
			<table border="0" width="760" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td colspan="2"><!--#INCLUDE FILE="admin_header.asp" --></td>
				</tr>
				<tr>
					<td width="187" valign="top" background="../images/bg_menuleft.gif"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="573" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="4" background="../images/bg_title.gif" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px">
								<b>&nbsp; THÊM MỚI THĂM DÒ</b></td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="4">&nbsp;</td>
							</tr>
							<form method="POST" name="frmAddNew" action="admin_addvote.asp">
							<tr>
								<td width="89">&nbsp;</td>
								<td width="111"><b>Tên câu hỏi</b></td>
								<td width="364">
								<input type="text" name="txtQuestion" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
							<tr>
								<td height="14" width="89"></td>
								<td height="14" width="111"><b>Trạng thái</b></td>
								<td height="14" width="364">
								<select size="1" name="txtStatus" class="input_text">
								<option selected value="1">Hiển thị</option>
								<option value="0">Ẩn</option>
								</select></td>
								<td height="14" width="9">
								</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Phương án 1</b></td>
								<td width="364">
								<input type="text" name="txtAnswer" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Phương án 2</b></td>
								<td width="364">
								<input type="text" name="txtAnswer" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Phương án 3</b></td>
								<td width="364">
								<input type="text" name="txtAnswer" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Phương án 4</b></td>
								<td width="364">
								<input type="text" name="txtAnswer" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
							</tr>
							<tr>
								<td width="89">
								&nbsp;</td>
								<td width="111">
								<b>Phương án 5</b></td>
								<td width="364">
								<input type="text" name="txtAnswer" size="28" class="input_text"></td>
								<td width="9">
								&nbsp;</td>
							</tr>
							<tr>
								<td width="573" colspan="4" height="20">
								<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
								</tr>
								<tr>
								<td width="89">&nbsp;</td>
								<td width="475" colspan="2">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
									<tr>
										<td width="145">&nbsp;</td>
										<td>
								<p align="left">
								<input type="submit" value="Tạo mới" name="B2" class="input_button">&nbsp;
								<input type="reset" value="Hủy bỏ" name="B3" class="input_button"></td>
									</tr>
								</table>
								</td>
								<td width="9">
								&nbsp;</td>
								</tr>
								<input type="hidden" name="category" value="vote">
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