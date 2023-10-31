<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<form method="POST" name="frmSearchBook" action="default.asp?name=advancebook">
				<td>
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						&nbsp;</td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						<p align="center" style="margin-top: 8px; margin-bottom: 10px">
						<font color="#003C5E"><b><font size="2">TÌM KIẾM NÂNG CAO</font><br>
						<img border="0" src="../images/line.gif" width="175" height="5"></b></font></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						<p align="center" style="margin-bottom: 6px; margin-top:10px"><b>
						<font size="2">Nhập từ khóa và tùy chọn tìm kiếm</font></b></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						<p align="center" style="margin-bottom: 4px">
						<font size="2">
						<input type="text" name="txtSearchKey" size="27" class="textbox"></font></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4" height="3"></td>
						<td width="98%" colspan="7" height="3">
						</td>
						<td width="1%" height="3"></td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="23%">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						&nbsp;</td>
						<td width="5%">
						<input type="checkbox" name="txtBookName" value="ON"></td>
						<td width="15%">
						<font size="2">Tên sách</font></td>
						<td width="5%">
						<input type="checkbox" name="txtSummary" value="ON"></td>
						<td width="13%">
						<font size="2">Tóm tắt</font></td>
						<td width="5%">
						<input type="checkbox" name="txtAuthorName" value="ON"></td>
						<td width="33%">
						<font size="2">Tác giả</font></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						&nbsp;</td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						<img border="0" src="../images/line.gif" width="250" height="3"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						<p align="center" style="margin-top: 6px; margin-bottom: 10px">
								<input type="submit" value="Tìm kiếm" name="B2" class="input_button">&nbsp;
								<input type="reset" value=" Hủy bỏ " name="B3" class="input_button"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="4">&nbsp;</td>
						<td width="98%" colspan="7">
						&nbsp;</td>
						<td width="1%">&nbsp;</td>
					</tr>
				</table>
				</td>
				<input type="hidden" name="category" value="searchbook">
				</form>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="5"></td>
	</tr>
</table>

</div>
</body>

</html>