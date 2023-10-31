<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
'<%
'	If Session("Username")= "" Then
'		Response.Redirect("admin_login.asp")
'	End If	
'%>
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
					<td width="187" valign="top"><!--#INCLUDE FILE="admin_menu.asp" --></td>
					<td width ="797" valign="top">
					<div align="center">
						<table border="0" width="573" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="9" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2" color="#FF0000">&nbsp;QUẢN LÝ SÁCH</font></b></td>
							</tr>
							<tr>
								<td colspan="9">&nbsp;</td>
							</tr>
							<tr>
								<td width="78" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="98" height="33">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="9" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="102" height="33">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="104" height="33">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="90" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="63" height="33">
								<p style="margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">
								<p align="center" style="margin-top: 0; margin-bottom: 3px">
								<a href="admin_genrecategory.asp">
								<font size="2">
								<img border="0" src="../images/category.gif" width="40" height="39"></font></a></td>
								<td width="9">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="102">
								<p style="margin-top: 0; margin-bottom: 3px" align="center">
								<a href="admin_bookcategory.asp">
								<font size="2">
								<img border="0" src="../images/genre.gif" width="40" height="39"></font></a></td>
								<td width="10">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="104">
								<p style="margin-top: 0; margin-bottom: 3px" align="center">
								<a href="admin_publisher.asp">
								<font size="2">
								<img border="0" src="../images/publisher.gif" width="40" height="39"></font></a></td>
								<td width="10">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="90">
								<p align="center" style="margin-top: 0; margin-bottom: 3px">
								<a href="admin_language.asp">
								<font size="2">
								<img border="0" src="../images/card.gif" width="40" height="37"></font></a></td>
								<td width="63">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">
								<p align="center" style="margin-top: 3px">
								<font size="2">Thể 
								loại sách</font></td>
								<td width="9">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="102">
								<p align="center" style="margin-top: 3px">
								<font size="2">Lĩnh vực</font></td>
								<td width="10">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="104">
								<p align="center" style="margin-top: 3px">
								<font size="2">Nhà 
								xuất bản</font></td>
								<td width="10">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="90">
								<p align="center" style="margin-top: 3px">
								<font size="2">Ngôn ngữ</font></td>
								<td width="63">
								<p style="margin-top: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="98" height="23">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="9" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="102" height="23">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="104" height="23">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="90" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="63" height="23">
								<p style="margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="9">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="102">
								<p align="center" style="margin-top: 0; margin-bottom: 3px">
								<a href="admin_addbook.asp">
								<font size="2">
								<img border="0" src="../images/add.gif" width="40" height="39"></font></a></td>
								<td width="10">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="104">
								<p style="margin-top: 0; margin-bottom: 3px" align="center">
								<a href="admin_listbook.asp">
								<font size="2">
								<img border="0" src="../images/list.gif" width="40" height="39"></font></a></td>
								<td width="10">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="90">
								<p align="left" style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-top: 0; margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="9">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="102">
								<p align="center" style="margin-top: 3px">
								<font size="2">Thêm sách mới</font></td>
								<td width="10">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="104">
								<p align="center" style="margin-top: 3px">
								<font size="2">Liệt kê - Sửa đổi</font></td>
								<td width="10">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="90">
								<p style="margin-top: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-top: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="98">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="9">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="102">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="104">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="90">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="102">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="104">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="90">
								<p align="left">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="102">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="104">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="90">
								<p align="left">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="98">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="9">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="102">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="104">
								<p style="margin-bottom: 3px" align="center">&nbsp;</td>
								<td width="10">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="90">
								<p style="margin-bottom: 3px">&nbsp;</td>
								<td width="63">
								<p style="margin-bottom: 3px">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="102">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="104">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="90">
								<p align="left">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td width="78">&nbsp;</td>
								<td width="98">&nbsp;</td>
								<td width="9">&nbsp;</td>
								<td width="102">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="104">&nbsp;</td>
								<td width="10">&nbsp;</td>
								<td width="90">
								<p align="left">&nbsp;</td>
								<td width="63">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="9">&nbsp;</td>
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
