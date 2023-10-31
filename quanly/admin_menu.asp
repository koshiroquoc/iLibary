<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body>

<table border="0" width="187" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="2" height="19" background="../images/bg_menu.jpg" align="center">
		<p><font size="2" color="#FF0000">&nbsp;</font><font color="#FF0000"><b><font size="2">MENU CHÍNH</font></b></font></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<font size="2"><font color="#000080">
		<%
			If Session("library") = True Then
		%>
		</font>
		<a href="admin_library.asp"><font color="#000080">Trang chủ</font></a><font color="#000080">
		<%
			Else
		%>
		</font>
		<a href="admin_default.asp"><font color="#000080">Trang chủ</font></a><font color="#000080">
		<%
			End If
		%>
		</font></font>
		</td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px"><a href="admin_news.asp">
		<font size="2" color="#000080">Quản lý tin tức</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px"><a href="admin_book.asp">
		<font size="2" color="#000080">Quản lý sách</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_document.asp"><font size="2" color="#000080">Quản lý tài liệu</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_libraring.asp"><font size="2" color="#000080">Quản lý mượn trả</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px"><a href="admin_card.asp">
		<font size="2" color="#000080">Quản lý thẻ</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_notice.asp"><font size="2" color="#000080">Quản lý thông báo</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_schedule.asp"><font size="2" color="#000080">Quản lý lịch trực</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 2px; margin-bottom: 2px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_software.asp"><font size="2" color="#000080">Quản lý tiện ích</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_relax.asp"><font size="2" color="#000080">Quản lý giải trí</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_listcontact.asp"><font size="2" color="#000080">Quản lý 
		góp ý</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px"><a href="admin_user.asp">
		<font size="2" color="#000080">Quản lý người dùng</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_setting.asp"><font size="2" color="#000080">Cấu hình website</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_changepass.asp"><font size="2" color="#000080">Thông tin cá nhân</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
	<tr>
		<td width="31" align="center">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<img border="0" src="../images/bullet.gif" width="9" height="9"></td>
		<td width="156" background="../images/bg_submenu.jpg">
		<p style="margin-top: 4px; margin-bottom: 4px">
		<a href="admin_logout.asp"><font size="2" color="#000080">Thoát</font></a></td>
	</tr>
	<tr>
		<td width="187" align="center" colspan="2" height="1">		
		<img border="0" src="../images/line_menu.gif" width="187" height="1"></td>
	</tr>
</table>
</body>

</html>