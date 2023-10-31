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
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="19">&nbsp;</td>
						<td>
						<p align="center">&nbsp;</td>
						<td width="20">&nbsp;</td>
					</tr>
					<tr>
						<td width="19">&nbsp;</td>
						<td>
						<p align="center" style="margin-top: 8px; margin-bottom: 10px">
						<font color="#003C5E">
						<b>TÌM KIẾM THEO TỰA ĐỀ<br>
						<img border="0" src="../images/line.gif" width="175" height="5"></b></font></td>
						<td width="20">&nbsp;</td>
					</tr>
					<tr>
						<td width="19">&nbsp;</td>
						<td>
						<p align="center" style="margin: 3px 4px"><b>
						<a href="default.asp?name=charbook&char=A">A</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=Ă">Ă</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=Â">Â</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=B">B</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=C">C</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=D">D</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=Đ">Đ</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=E">E</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=Ê">Ê</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=F">F</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=G">G</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=H">H</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=I">I</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=J">J</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=K">K</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=L">L</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=M">M</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=N">N</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=O">O</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=Ô">Ô</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=Ơ">Ơ</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=P">P</a></b></p>
						<p align="center" style="margin: 3px 4px"><b>
						<a href="default.asp?name=charbook&char=Q">Q</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=R">R</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=S">S</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=T">T</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=U">U</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=Ư">Ư</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=V">V</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=W">W</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=X">X</a>&nbsp;&nbsp; 
						<a href="default.asp?name=charbook&char=Y">Y</a>&nbsp;&nbsp;
						<a href="default.asp?name=charbook&char=Z">Z</a></b></td>
						<td width="20">&nbsp;</td>
					</tr>
					<tr>
						<td width="19">&nbsp;</td>
						<td>
						<p align="center" style="margin-top: 10px; margin-bottom: 10px">
						Kích chọn chữ cái bắt đầu của tựa đề.</td>
						<td width="20">&nbsp;</td>
					</tr>
				</table>
				</td>
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
				<form method="POST" name="frmSearchBook" action="default.asp?name=bookresult">
				<td>
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						&nbsp;</td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						<p align="center" style="margin-top: 8px; margin-bottom: 10px">
						<font color="#003C5E"><b>TÌM KIẾM THEO TÊN SÁCH<br>
						<img border="0" src="../images/line.gif" width="175" height="5"></b></font></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						<p align="center" style="margin-bottom: 6px; margin-top:10px"><b>
						Nhập từ khóa và tùy chọn tìm kiếm</b></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						<p align="center" style="margin-bottom: 4px">
						<input type="text" name="txtSearchKey" size="27" class="textbox"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						<p align="center" style="margin-top: 6px; margin-bottom: 6px">
						<img border="0" src="../images/line.gif" width="250" height="3"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
						<p align="center" style="margin-top: 6px; margin-bottom: 10px">
								<input type="submit" value="Tìm kiếm" name="B2" class="input_button">&nbsp;
								<input type="reset" value=" Hủy bỏ " name="B3" class="input_button"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td colspan="2">
						<p align="right" style="margin-right: 10px"><b>
						<a href="default.asp?name=advancesearch">Tìm kiếm nâng 
						cao</a></b></td>
					</tr>
					<tr>
						<td width="13">&nbsp;</td>
						<td width="96%">
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