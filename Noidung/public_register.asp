<%	Session.CodePage = 65001 %>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Đăng ký mượn sách</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
<script language="javascript">
	function CheckInput(){		
		if(document.frmRegister.txtCardID.value == ""){
			alert("Bạn chưa nhập mã thẻ thư viện!");
			document.frmRegister.txtTitle.focus();
			return;
			}
		if(document.frmRegister.txtFullname.value == ""){
			alert("Họ tên chưa nhập họ và tên!");
			document.frmRegister.txtFullname.focus();
			return;
		}
		if(document.frmRegister.txtFullname.value.length <6){
			alert("Họ và tên không hợp lệ!");
			document.frmRegister.txtFullname.focus();
			return;
		}
		document.frmRegister.submit();
	}
</script>
</head>

<body topmargin="5" leftmargin="5">
<div align="center">
<table border="0" width="431" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
			<!-- #INCLUDE FILE="../include/inc_function.asp" -->
			<%
				id	= Request.QueryString("id")
				If id <> "" Then
				Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
				strSQL = "SELECT ID, BOOK_ID, NAME FROM BOOK WHERE ID=" & id
				rsCheckBook.Open strSQL, Conn,3,1
				If Not rsCheckBook.Eof Then
					txtBookID = rsCheckBook("BOOK_ID")
					txtName = rsCheckBook("NAME")
					rsCheckBook.Close
					Set rsCheckBook = Nothing
				End If
				End If
				txtCategory	= Request.Form("category")	
				If txtCategory = "register" Then				
					txtCardID	= Request.Form("txtCardID")
					txtBookID	= Request.Form("txtBookID")
					txtFullname	= Request.Form("txtFullname")
					
					If CheckCardID(txtCardID) = 0 Then
						txtFail2 = True
					ElseIf CountCardID(txtCardID)>=3 Then
						txtFail = True
					ElseIf CountRegister(txtCardID,txtBookID)>=1 Then
						txtFail1 = True
					Else		
						strSQL = "SELECT * FROM REGISTER Order by ID Desc"		
						txtID = GetID(strSQL,Conn)
						
						strSQL = "INSERT INTO REGISTER(ID,CARD_ID,BOOK_ID,FULLNAME,DATE_INFORM) VALUES("
						strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
						strSQL = strSQL & CheckString(txtBookID,",")& CheckString(txtFullName,",")& CheckString(Now(),")")
						Conn.Execute strSQL
						Conn.Close
						Set Conn = Nothing
					End If	
			%>
			<%
				If txtFail1 = True Then
			%>
			<tr>
				<td>
				<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="53">&nbsp;</td>
						<td>
						<table border="1" width="100%" id="table3" bordercolorlight="#F5F5F5" bordercolordark="#CECECE" cellspacing="0" cellpadding="0">
							<tr>
								<td><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/iconerror.gif" width="46" height="44"></p>
													<p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<b>Lổi! &nbsp;Bạn đã đăng ký 1 mã sách 2 lần liên tiếp nhau!</b></p>
													<p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/line.gif" width="130" height="5"></p>
													<p align="center" style="margin-top: 6px; margin-bottom: 6px">
									<a href="JavaScript:history.back();">Quay lại</a></p>
													<p align="center" style="margin-top: 13px; margin-bottom: 6px">
									<img border="0" src="../images/line.gif" width="269" height="5"></p>
													<p align="center" style="margin-top: 6px; margin-bottom: 6px">
									Website hỗ trợ thông tin</p>
								</td>
							</tr>
						</table>
						</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				ElseIf txtFail = True Then
			%>
			<tr>
				<td>
				<table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="53">&nbsp;</td>
						<td>
						<table border="1" width="100%" id="table7" bordercolorlight="#F5F5F5" bordercolordark="#CECECE" cellspacing="0" cellpadding="0">
							<tr>
								<td><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/iconerror.gif" width="46" height="44"></p><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<b>Xin lổi! Bạn chỉ được đăng ký tối đa là 3 lần!</b></p>
									<p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/line.gif" width="130" height="5"><p align="center" style="margin-top: 4px; margin-bottom: 4px">
									<a href="#" onclick="JavaScript:window.close()">Đóng</a><p align="center" style="margin-top: 13px; margin-bottom: 6px">
									<img border="0" src="../images/line.gif" width="269" height="5"></p>
													<p align="center" style="margin-top: 6px; margin-bottom: 6px">
									Website hỗ trợ thông tin</p>
								</td>
							</tr>
						</table>
						</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				ElseIf txtFail2 = True Then
			%>
			<tr>
				<td>
				<table border="0" width="100%" id="table10" cellspacing="0" cellpadding="0">
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="53">&nbsp;</td>
						<td>
						<table border="1" width="100%" id="table11" bordercolorlight="#F5F5F5" bordercolordark="#CECECE" cellspacing="0" cellpadding="0">
							<tr>
								<td><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/iconerror.gif" width="46" height="44"></p><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<b>Mã thẻ không hợp lệ hoặc không tồn tại.</b></p>
													<p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<b>Xin vui lòng xem lại!</b></p>
									<p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/line.gif" width="130" height="5"><p align="center" style="margin-top: 4px; margin-bottom: 4px">
									<a href="JavaScript:history.back();">Quay lại</a><p align="center" style="margin-top: 13px; margin-bottom: 6px">
									<img border="0" src="../images/line.gif" width="269" height="5"></p>
													<p align="center" style="margin-top: 6px; margin-bottom: 6px">
									Website hỗ trợ thông tin</p>
								</td>
							</tr>
						</table>
						</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				Else
			%>
			<tr>
				<td>
				<table border="0" width="100%" id="table14" cellspacing="0" cellpadding="0">
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="53">&nbsp;</td>
						<td>
						<table border="1" width="100%" id="table15" bordercolorlight="#F5F5F5" bordercolordark="#CECECE" cellspacing="0" cellpadding="0">
							<tr>
								<td><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/Pic/saveitem.gif" width="16" height="16"><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<b>Thông tin đăng ký đã được cập nhật<br>
									Bạn hãy đến thư viện để liên hệ lấy sách.</b><p align="center" style="margin-top: 10px; margin-bottom: 10px">
									<img border="0" src="../images/line.gif" width="130" height="5"><p align="center" style="margin-top: 4px; margin-bottom: 4px">
									<a href="#" onclick="JavaScript:window.close()">Đóng</a><p align="center" style="margin-top: 13px; margin-bottom: 6px">
									<img border="0" src="../images/line.gif" width="269" height="5"></p>
													<p align="center" style="margin-top: 6px; margin-bottom: 6px">
									Website hỗ trợ thông tin</p>
								</td>
							</tr>
						</table>
						</td>
						<td width="53">&nbsp;</td>
					</tr>
					<tr>
						<td width="42">&nbsp;</td>
						<td>&nbsp;</td>
						<td width="53">&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				End If
			%>
			<%
				Else			
			%>
			<tr>			
				<form method="POST" name="frmRegister" action="public_register.asp">
				<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%">&nbsp;</td>
				<td width="79%" colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="4" align="center"><b>ĐĂNG KÝ MƯỢN SÁCH</b></td>
			</tr>
			<tr>
				<td width="5%" height="6"></td>
				<td width="16%" height="6"></td>
				<td width="53%" height="6"></td>
				<td width="41%" height="6"></td>
			</tr>
			<tr>
				<td width="117%" colspan="4">
				<p align="center" style="margin-top: 0; margin-bottom: 0"><i>Khi đăng ký 
				thành công, sách sẽ được dành riêng cho 
				bạn,</i><p align="center" style="margin-top: 0; margin-bottom: 0">
				<i>&nbsp;xin liên hệ với Thư viện để nhận.<br>
				<font color="#000080">Sau 3 ngày nếu bạn không liên hệ mượn thì 
				hệ thống sẽ xoá đăng ký.</font></i></td>
				</tr>
			<tr>
				<td width="115%" height="6" colspan="4">
				<p align="center" style="margin-top: 4px; margin-bottom: 4px">
				<img border="0" src="../images/line.gif" width="226" height="5"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%"><b>Mã sách</b></td>
				<td width="79%" colspan="2">
				<font color="#006600"><b><%=txtBookID%></b></font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="16%" height="4"></td>
				<td height="4" width="79%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%"><b>Tên sách</b></td>
				<td width="79%" colspan="2">
				<b><font color="#006600"><%=txtName%></font></b></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="16%" height="4"></td>
				<td height="4" width="79%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%"><b>Mã thẻ</b></td>
				<td width="79%" colspan="2">
				<input type="text" name="txtCardID" size="11" class="textbox">
				<font color="#FF0000">&nbsp;*</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="16%" height="4"></td>
				<td height="4" width="79%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%"><b>Họ tên</b></td>
				<td width="79%" colspan="2">
				<input type="text" name="txtFullname" size="52" class="textbox"><font color="#FF0000">&nbsp; 
				*</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="16%" height="4"></td>
				<td height="4" width="79%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%">&nbsp;</td>
				<td width="53%">
				&nbsp;</td>
				<td width="16%">
				&nbsp;</td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="16%" height="4"></td>
				<td width="79%" height="4" colspan="2"></td>
			</tr>
			<tr>
				<td width="99%" colspan="4">
				<p align="center">
				<button name="B1" class="input_button" onclick="JavaScript:CheckInput();">&nbsp; 
				Đăng ký &nbsp;
				</button></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="16%">&nbsp;</td>
				<td width="79%" colspan="2">&nbsp;</td>
			</tr>
		</table>
					</td>
					<input type="hidden" name="category" value="register">
				<input type="hidden" name="txtBookID" value="<%=txtBookID%>">
				</form>
			</tr>
		<%
			End If
		%>	
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
	</table>

</div>
</body>

</html>