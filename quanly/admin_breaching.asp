<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("book") = False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If
	End If	
%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body topmargin="8" leftmargin="8">
<%
	txtClassID = Request.Form("txtClassID")
	If txtClassID = "" Then
		txtClassID = "All"
	End If	
%>
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
					<td width ="797" height ="350" valign = "top">
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">&nbsp;<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<b>&nbsp;<font size="2" color="#FF0000">ĐỘC GIẢ MƯỢN SÁCH QUÁ HẠN</font></b></td>
							</tr>
						<tr>
							<td width="10">&nbsp;</td>
							<td>&nbsp;</td>
							<td width="10">&nbsp;</td>
						</tr>
						<form method="POST" name="frmList">
						<tr>
							<td width="10">&nbsp;</td>
							<td>
							<table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC" height="20" bgcolor="#EEEEEE">
								<tr>
									<td>
									<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0" height="21">
										<tr>
											<td width="120"><b>&nbsp; Lọc theo 
											loại thẻ:</b></td>
											<td>
											<p style="margin-top: 1px">
											<select size="1" name="txtClass" class="input_text" onchange="JavaScript:cboChangeClass('txtClass');">
											<option selected value="All">-- Tất cả --</option>
											<%
											strSQL = "Select DISTINCT CATEGORY_ID, name From CATEGORY_CARD"
												Call ListComboCARD1(strSQL,Request.Form("txtCategoryFilter"))
											%>
											</select></td>
											<td width="44">
											<table border="1" width="65" id="table11" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onClick="JavaScript:doSubmit('admin_delete.asp?category=breach');"><p align="center">
													<span style="font-size: 8pt">
													Xóa</span></td>
												</tr>
											</table>
											</td>
											<td width="4">&nbsp;</td>
											<td width="76" valign="middle">
											<table border="1" width="75" id="table10" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onClick="JavaScript:openWindowPrint('admin_print.asp?typeprint=breaching&class=<%=txtClassID%>');"><p align="center">
													<span style="font-size: 8pt">
													In danh sách</span></td>
												</tr>
											</table>
											</td>
											<td width="3"></td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
							</td>
							<td width="10">&nbsp;</td>
						</tr>
						<tr>
							<td width="10">&nbsp;</td>
							<td>&nbsp;</td>
							<td width="10">&nbsp;</td>
						</tr>
						<tr>
							<td width="10">&nbsp;</td>
							<td>
							<%
								txtCategoryFilter = Request.Form("txtCategoryFilter")
								If txtCategoryFilter ="" Then
									strSQL = "SELECT * FROM BORROW WHERE NOW()-DATE_INFORM>7"
									strSQL =strSQL & " Order by BOOK_ID ASC"
								Else
									If txtCategoryFilter ="All" Then
										strSQL = "SELECT * FROM BORROW WHERE NOW()-DATE_INFORM>7"
										strSQL =strSQL & " Order by BOOK_ID ASC"
									Else
										strSQL = "SELECT * FROM BORROW WHERE left(CARD_ID,2) ='"& txtCategoryFilter & "'"
										strSQL =strSQL & " AND NOW()-DATE_INFORM>7"
										strSQL = strSQL & " Order by CARD_ID ASC"
									End If								
								End If	
								
								Set rsSelect = Server.CreateObject("ADODB.Recordset")
								rsSelect.Open strSQL,Conn,3,1

							  	If Request.QueryString("page") = "" Then
									intCurrentPage = 1
								Else
									intCurrentPage = CInt(Request.QueryString("page"))
								End If
								rsSelect.PageSize = 15
								If rsSelect.PageCount > 0 then
									rsSelect.AbsolutePage = intCurrentPage
								Else
									intCurrentPage = 0
								End If
								
								numPage = rsSelect.PageCount
								If Not rsSelect.Eof Then
							%>
							<table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
								<tr>
									<td width="26" align="center" bgcolor="#EEEEEE" height="20">&nbsp;</td>
									<td width="26" align="center" bgcolor="#EEEEEE" height="20"><b>STT</b></td>
									<td align="center" width="57" bgcolor="#EEEEEE" height="20">
									<b>Mã thẻ</b></td>
									<td align="center" width="128" bgcolor="#EEEEEE" height="20">
									<b>Họ và tên</b></td>
									<td align="center" width="56" bgcolor="#EEEEEE" height="20">
									<b>Đơn vị</b></td>
									<td align="center" width="166" bgcolor="#EEEEEE" height="20">
									<b>Tên sách</b></td>
									<td align="center" bgcolor="#EEEEEE" height="20">
									<b>Ngày mượn</b></td>
								</tr>
								<%
									Dim iCount
									iCount = 1
									Do While Not rsSelect.Eof and iCount <=rsSelect.PageSize
									strSQL = "SELECT BOOK_ID, NAME FROM BOOK WHERE BOOK_ID='" & rsSelect("BOOK_ID") & "'"
									Set rsCategory = Server.CreateObject("ADODB.Recordset")
									rsCategory.Open strSQL,Conn,3,1
									strSQL = "SELECT * FROM CARD WHERE CARD_ID='" & rsSelect("CARD_ID") & "'"
									Set rsName = Server.CreateObject("ADODB.Recordset")
									rsName.Open strSQL,Conn,3,1
								%>
								<tr>
									<td width="26" align="center">
									<input type="CHECKBOX" name="Mid" value="<%=rsSelect("ID")%>"></td>
									<td width="26" align="center"><%=iCount%></td>
									<td width="56">
									<p align="center" style="margin-left: 4px; margin-right: 4px">
									<%=rsSelect("CARD_ID")%></td>
									<td width="128">
									<p align="left" style="margin-left: 4px; margin-right: 4px">
									<%=rsName("FIRSTNAME") & " " & rsName("LASTNAME")%></td>
									<td width="56">
									<p align="center" style="margin-left: 4px; margin-right: 4px">
									<%=rsSelect("CLASS_ID")%></td>
									<td width="166">
									<p align="justify" style="margin-left: 4px; margin-right: 4px"><%=rsCategory("NAME")%></td>
									<td align ="center"><%=NgayVN(rsSelect("DATE_INFORM"))%></td>
								</tr>
								<%
									iCount = iCount + 1
									rsSelect.MoveNext
									Loop
									Conn.Close
									Set Conn = Nothing									
								%>
							</table>
							<%
								End If
							%>
							</td>
							<td width="10">&nbsp;</td>
						</tr>
						</form>
						<tr>
							<td width="10">&nbsp;</td>
							<td>&nbsp;</td>
							<td width="10">&nbsp;</td>
						</tr>
						<tr>
							<td width="10">&nbsp;</td>
							<td>
							<%
								If numPage >1 Then
							%>
							<p align="right">Trang: 
							<%
								for i=1 to numPage
							%>	
								<b>	
								<a href="?page=<%=i%>"><%=i%></a> |</b>
							<%	
								Next
								End If
							%>
							</td>
							<td width="10">&nbsp;</td>
						</tr>
						<form method="POST" name="frmFilter" action="admin_breaching.asp">
							<input type="hidden" name="txtCategoryFilter" value="">
							<input type="hidden" name="txtClassID" value="">
						</form>						
						<tr>
							<td width="10">&nbsp;</td>
							<td>&nbsp;</td>
							<td width="10">&nbsp;</td>
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