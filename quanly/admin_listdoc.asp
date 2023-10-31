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
					<td width ="797" height ="170" valign = "top">
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="3" height="19">
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<font color="#FF0000" size="2"><b>&nbsp; 
							DANH SÁCH TÀI LIỆU</b></font></td>
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
											<td width="147">&nbsp;<b> Lọc theo 
											lĩnh vực tài liệu</b></td>
											<td>
											<p style="margin-top: 2px">
											<select size="1" name="txtCategory" class="input_text" onchange="JavaScript:cboChange('txtCategory');">
											<option selected value="All">-- Tất cả --</option>
											<%
												strSQL = "Select NAME, CATEGORY_ID From CATEGORY_DOCUMENT"
												Call ListCombo(strSQL,Request.Form("txtCategoryFilter"))
											%>
											</select></td>
											<td width="65" valign="middle">
											<table border="1" width="65" id="table8" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onclick="JavaScript:doSubmit('admin_delete.asp?category=document');"><p align="center">
													<span style="font-size: 8pt">
													Xóa</span></td>
												</tr>
											</table>
											</td>
											<td width="4">&nbsp;</td>
											<td width="56" valign="middle">
											<table border="1" width="60" id="table10" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onClick="window.location.href='admin_adddoc.asp'"><p align="center">
													<span style="font-size: 8pt">
													Tạo mới</span></td>
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
									strSQL = "SELECT * FROM DOCUMENT Order by DATE_INFORM Desc"
								Else
									If txtCategoryFilter ="All" Then
										strSQL = "SELECT * FROM DOCUMENT Order by DATE_INFORM Desc"
									Else
										strSQL = "SELECT * FROM DOCUMENT WHERE LEFT(DOCUMENT_ID,3)='"& txtCategoryFilter & "'"
										strSQL = strSQL & "Order by DATE_INFORM Desc"
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
									<td align="center" width="202" bgcolor="#EEEEEE" height="20">
									<b>Tên tài liệu</b></td>
									<td align="center" width="100" bgcolor="#EEEEEE" height="20">
									<b>Thể loại</b></td>
									<td align="center" bgcolor="#EEEEEE" height="20">
									<b>Tác giả</b></td>
									<td align="center" width="61" bgcolor="#EEEEEE" height="20">
									<b>Ngày đưa</b></td>
								</tr>
								<%
									Dim iCount
									iCount = 1
									Do While Not rsSelect.Eof and iCount <=rsSelect.PageSize
									strSQL = "SELECT * FROM CATEGORY_DOCUMENT WHERE CATEGORY_ID='" & Left(rsSelect("DOCUMENT_ID"),3) & "'"
									Set rsCategory = Server.CreateObject("ADODB.Recordset")
									rsCategory.Open strSQL,Conn,3,1

								%>
								<tr>
									<td width="26" align="center">
									<input type="CHECKBOX" name="Mid" value="<%=rsSelect("ID")%>"></td>
									<td width="26" align="center"><%=iCount%></td>
									<td width="202">
									<p align="justify" style="margin-left: 4px; margin-right: 4px">
									<a href="admin_editdoc.asp?id=<%=rsSelect("ID")%>"><%=rsSelect("TITLE")%></a></td>
									<td width="100">
									<p align="justify" style="margin-left: 4px; margin-right: 4px"><%=rsCategory("NAME")%></td>
									<td>
									<p align="justify" style="margin-left: 4px; margin-right: 4px"><%=rsSelect("AUTHOR")%></td></td>
									<td width="68" align="center"><%=NgayVN(rsSelect("DATE_INFORM"))%></td>
								</tr>
								<%
									iCount = iCount + 1
									rsSelect.MoveNext
									Loop
									rsCategory.Close
									Set rsCategory = Nothing
									rsSelect.Close
									Set rsSelect = Nothing
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
						<form method="POST" name="frmFilter" action="admin_listdoc.asp">
							<input type="hidden" name="txtCategoryFilter" value="">
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