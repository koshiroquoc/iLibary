<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>
<!-- #INCLUDE FILE="../include/inc_js.asp" -->
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
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="3" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2" color="#FF0000">&nbsp;</font></b><p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="3" color="#FF0000">QUẢN LÝ THẺ THƯ VIỆN</font></b></td>
							</tr>
							<tr>
								<td width="12">&nbsp;</td>
								<td width="548">
								<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" cellspacing="0" cellpadding="0">
											<tr>
												<td width="10" height="35">&nbsp;</td>
												<td width="120" height="35">&nbsp;</td>
												<td width="9" height="35">&nbsp;</td>
												<td width="110" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="28" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="112" height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
												<td height="35">&nbsp;</td>
												<td width="10" height="35">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="120">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_cardcategory.asp" onMouseOver="JavaScript:handleOver('card_cate','card_cate1');return true;" onMouseOut="JavaScript:handleOut('card_cate','card_cate');return true;">
												<img border="0" name = "card_cate" src="../images/card_cate.gif" width="40" height="39"></a></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_importcard.asp" onMouseOver="JavaScript:handleOver('card_excel','card_excel1');return true;" onMouseOut="JavaScript:handleOut('card_excel','card_excel');return true;">
												<img border="0" name="card_excel" src="../images/card_excel.gif" width="40" height="39"></a></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_createcard.asp" onMouseOver="JavaScript:handleOver('card_new','card_new1');return true;" onMouseOut="JavaScript:handleOut('card_new','card_new');return true;">
												<img border="0" name="card_new" src="../images/card_new.gif" width="40" height="39"></a></td>
												<td>&nbsp;</td>
												<td>
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												<a href="admin_listcard.asp" onMouseOver="JavaScript:handleOver('list','list1');return true;" onMouseOut="JavaScript:handleOut('list','list');return true;">
												<img border="0" name="list" src="../images/list.gif" width="40" height="39"></a></td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="120">
												<p align="center"><b>Danh mục 
												độc giả</b></td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center"><b>Tạo thẻ từ 
												Excel</b></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center"><b>Bổ sung thẻ</b></td>
												<td>&nbsp;</td>
												<td>
												<p align="center"><b>
								Liệt kê - Sửa đổi</b></td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="120" height="10"></td>
												<td width="9" height="10"></td>
												<td height="10"></td>
												<td width="8" height="10"></td>
												<td height="10"></td>
												<td width="13" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											<tr>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="120">
												<p align="center" style="margin-top: 0; margin-bottom: 5px">
												&nbsp;</td>
												<td width="9">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_listbreach.asp" onMouseOver="JavaScript:handleOver('breach','breach1');return true;" onMouseOut="JavaScript:handleOut('breach','breach');return true;">
												<img border="0" name="breach" src="../images/breach.gif" width="48" height="47"></a></td>
												<td width="8">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td width="13">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												<a href="admin_breaching.asp" onMouseOver="JavaScript:handleOver('breaching','breaching1');return true;" onMouseOut="JavaScript:handleOut('breaching','breaching');return true;">
												<img border="0" name ="breaching" src="../images/breaching.gif" width="40" height="39"></a></td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
												<td>
												<p style="margin-top: 0; margin-bottom: 5px" align="center">
												&nbsp;</td>
												<td width="10">
												<p style="margin-top: 0; margin-bottom: 5px">&nbsp;</td>
											</tr>
											<tr>
												<td width="13">&nbsp;</td>
												<td width="120">
												<p align="center">&nbsp;</td>
												<td width="9">&nbsp;</td>
												<td>
												<p align="center"><b>DS thẻ đang 
												phạt</b></td>
												<td width="8">&nbsp;</td>
												<td>&nbsp;</td>
												<td width="13">&nbsp;</td>
												<td>
												<p align="center"><b>DS thẻ vi 
												phạm</b></td>
												<td>&nbsp;</td>
												<td>
												<p align="center">&nbsp;</td>
												<td width="10">&nbsp;</td>
											</tr>
											<tr>
												<td width="13" height="10"></td>
												<td width="120" height="10"></td>
												<td width="9" height="10"></td>
												<td height="10"></td>
												<td width="8" height="10"></td>
												<td height="10"></td>
												<td width="13" height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td height="10"></td>
												<td width="10" height="10"></td>
											</tr>
											</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="13">&nbsp;</td>
							</tr>
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