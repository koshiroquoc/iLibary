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
							<p style="margin-top: 2px; margin-bottom: 2px" align="center"><b>&nbsp; 
							</b>
							<p style="margin-top: 2px; margin-bottom: 2px" align="center">
							<b><font color="#FF0000" style="font-size: 11pt">DANH SÁCH THẺ THƯ 
							VIỆN</font></b></td>
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
											<td width="116">&nbsp;<b> Lọc theo 
											loại thẻ</b></td>
											<td>
											<p style="margin-top: 2px">
											<select size="1" name="txtCategory" class="input_text" onchange="JavaScript:cboChange('txtCategory');">
											<option selected value="All">-- Tất cả --</option>
											<%
												strSQL = "Select DISTINCT CATEGORY_ID, name From CATEGORY_CARD"
												Call ListComboCARD1(strSQL,Request.Form("txtCategoryFilter"))
											%>
											</select></td>
											<td width="65" valign="middle">
											<table border="1" width="65" id="table8" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onclick="JavaScript:doSubmit('admin_delete.asp?category=card');"><p align="center">
													<span style="font-size: 8pt">
													Xóa</span></td>
												</tr>
											</table>
											</td>
											<td width="4">&nbsp;</td>
											<td width="56" valign="middle">
											<table border="1" width="60" id="table10" cellspacing="0" cellpadding="0" bgcolor="#EEEEEE" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" height="19">
												<tr>
													<td valign="top" style="cursor:hand" onMouseOver="this.style.background='#FFFFFF'" onMouseOut="this.style.background='#EEEEEE'" onClick="window.location.href='admin_createcard.asp'">
													<p align="center">
													<span style="font-size: 8pt">
													Bổ sung</span></td>
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
									strSQL = "SELECT * FROM CARD Order by CARD_ID Asc"
								Else
									If txtCategoryFilter ="All" Then
										strSQL = "SELECT * FROM CARD Order by CARD_ID Asc"
									Else
										strSQL = "SELECT * FROM CARD WHERE CATACARD_ID='"& txtCategoryFilter & "'"
										strSQL = strSQL & "Order by CARD_ID Asc"
									End If								
								End If	

								Set rsSelect = Server.CreateObject("ADODB.Recordset")
								rsSelect.Open strSQL,Conn,3,1

							  	If Request.QueryString("page") = "" Then
									intCurrentPage = 1
								Else
									intCurrentPage = CInt(Request.QueryString("page"))
								End If
								rsSelect.PageSize = 20
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
									<td width="31" align="center" bgcolor="#EEEEEE" height="20"><b>STT</b></td>
									<td align="center" width="70" bgcolor="#EEEEEE" height="20">
									<b>Loại thẻ</b></td>
									<td align="center" width="70" bgcolor="#EEEEEE" height="20">
									<b>Mã thẻ</b></td>
									<td align="center" width="162" bgcolor="#EEEEEE" height="20">
									<b>Họ và tên</b></td>
									<td align="center" bgcolor="#EEEEEE" height="20" width="90">
									<b>Ngày sinh</b></td>
									<td align="center" bgcolor="#EEEEEE" height="20">
									<b>Lớp</b></td>
									<td align="center" bgcolor="#EEEEEE" height="20">
									<b>Năm học</b></td>
									<td align="center" width="81" bgcolor="#EEEEEE" height="20">
									<b>Ngày tạo</b></td>
								</tr>
								<%
									Dim iCount
									iCount = 1
									Do While Not rsSelect.Eof and iCount <=rsSelect.PageSize
								%>
								<tr>
									<td width="26" align="center">
									<input type="CHECKBOX" name="Mid" value="<%=rsSelect("ID")%>"></td>
									<td width="31" align="center"><%=iCount%></td>
																		
									<td width="70">
									<p align="center" style="margin-left: 4px; margin-right: 4px">
									<%=rsSelect("CATACARD_ID")%></td>

									<td width="70">
									<p align="center" style="margin-left: 4px; margin-right: 4px">
									<%=rsSelect("CARD_ID")%> </td>
									<td width="162">
									<p align="left" style="margin-left: 4px; margin-right: 4px"><%=rsSelect("FIRSTNAME") & " " & rsSelect("LASTNAME")%></td>
									<td width="90">
									<p align="center" style="margin-left: 4px; margin-right: 4px"><%=NgayVN(rsSelect("BIRTHDAY"))%></td>
									<td>
									<p align="center" style="margin-left: 4px; margin-right: 4px"><%=rsSelect("CLASS_ID")%></td></td>
									<td>
									<p align="center" style="margin-left: 4px; margin-right: 4px"><%=rsSelect("SCHOOLYEAR")%></td></td>
									<td width="81" align="center"><%=NgayVN(rsSelect("DATE_INFORM"))%></td>
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
								<a href="?page=<%=i%>&"><%=i%></a> |</b>
							<%	
								Next
								End If
							%>
							</td>
							<td width="10">&nbsp;</td>
						</tr>
						<form method="POST" name="frmFilter" action="admin_listcard.asp">
							<input type="hidden" name="txtCategoryFilter" value="">
						</form>						
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