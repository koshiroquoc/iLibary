<%	Session.CodePage = 65001 %>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
<script language="javascript">
	function checkemail(email){
	  var etmp = email;
		for(var i=1; i < etmp.length-2; i++){	
			if( etmp.charAt(i) == '@'){
				for(var j=i+1; j <etmp.length -1; j++){
					if( etmp.charAt(j) == '.') 
					  return true;
				}
			}
		}
		return false;
	}
					
	function CheckInput(){		
		if(document.frmContact.txtTitle.value == ""){
			alert("Bạn chưa nhập chủ đề góp ý!");
			document.frmContact.txtTitle.focus();
			return;
			}
		if(document.frmContact.txtFullname.value == ""){
			alert("Bạn chưa nhập họ tên!");
			document.frmContact.txtFullname.focus();
			return;
		}		
		if(document.frmContact.txtEmail.value == ""){
			alert("Bạn chưa nhập địa chỉ email!");
			document.frmContact.txtEmail.focus();
			return;
		}
		if(!checkemail(document.frmContact.txtEmail.value)){
			alert("Địa chỉ email không hợp lệ, xin nhập lại!");
			document.frmContact.txtEmail.value = "";
			document.frmContact.txtEmail.focus();
			return;
		}
		if(document.frmContact.txtContent.value == ""){
			alert("Bạn chưa nhập nội dung góp ý!");
			document.frmContact.txtContent.focus();
			return;
		}
			document.frmContact.submit();
	}
</script>
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
			<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
			<!-- #INCLUDE FILE="../include/inc_function.asp" -->
			<%
				txtCategory	= Request.Form("category")	
				If txtCategory = "contact" Then				
					txtTitle	= Request.Form("txtTitle")
					txtEmail	= Request.Form("txtEmail")
					txtFullname	= Request.Form("txtFullname")
					txtContent	= Request.Form("txtContent")
					
					strSQL = "SELECT * FROM CONTACT Order by ID Desc"		
					txtID = GetID(strSQL,Conn)
					
					strSQL = "INSERT INTO CONTACT(ID,TITLE,EMAIL,FULLNAME,CONTENT,DATE_INFORM) VALUES("
					strSQL = strSQL & CheckString(txtID,",") & CheckString(txtTitle,",") & CheckString(txtEmail,",")
					strSQL = strSQL & CheckString(txtFullname,",") & CheckString(txtContent,",")& CheckString(Now(),")")
					Conn.Execute strSQL
					Conn.Close
					Set Conn = Nothing
			%>
			<tr>
				<td>
				<p align="center" style="margin-top: 10px; margin-bottom: 10px">
				&nbsp;<p align="center" style="margin-top: 10px; margin-bottom: 10px">
				<i><font size="2">Thông tin góp ý của bạn đã được gửi.&nbsp;
				Xin cảm ơn đã liên hệ với chúng tôi !</font></i><p align="center" style="margin-top: 10px; margin-bottom: 10px">
				<font size="3">
				<img border="0" src="../images/line.gif" width="130" height="5"></font><p align="center" style="margin-top: 10px; margin-bottom: 10px">
				<i><b>
				<font size="2">Website THƯ VIỆN TRỰC TUYẾN</font></b></i><font size="3"> </font></td>
			</tr>
			<%
				Else			
			%>
			<tr>			
				<form method="POST" name="frmContact" action="default.asp?name=contact">
				<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%">&nbsp;</td>
				<td width="80%" colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="4" align="center">
				<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
				<p style="margin-top: 0; margin-bottom: 0">
				<font color="#000080" size="2"><b>Trường THCS Trưng Vương, TP. Đà Nẵng<br>
&nbsp;</b></font></p>
				<p style="margin-top: 0; margin-bottom: 0">
				<font color="#003300" face="Verdana" size="2">Địa chỉ: 88 Yên Bái, TP.Đà Nẵng</font></p>
				<p style="margin-top: 0; margin-bottom: 0">
				<font color="#003300" face="Verdana" size="2">Website: 
				http://trungvuong.edu.vn/thuvien</font></p>
				<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
				<p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%">&nbsp;</td>
				<td width="64%">&nbsp;</td>
				<td width="41%">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%">&nbsp;</td>
				<td width="64%">
				<p align="center"><i><font size="2">Mục dấu (*) là bắt buộc</font></i></td>
				<td width="41%">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%"><font size="2" face="Verdana">Tiêu đề </font>
				<font size="2" color="#FF0000">
				(*)</font></td>
				<td width="80%" colspan="2">
				<font size="3">
				<input type="text" name="txtTitle" size="52" class="textbox">
				</font><font size="3" color="#FF0000">&nbsp;</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="15%" height="4"></td>
				<td height="4" width="80%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%"><font size="2" face="Verdana">Họ tên </font>
				<font color="#FF0000" size="2">
				(*)</font></td>
				<td width="80%" colspan="2">
				<font size="3">
				<input type="text" name="txtFullname" size="52" class="textbox"></font><font color="#FF0000" size="3">&nbsp; 
				</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="15%" height="4"></td>
				<td height="4" width="80%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="15%" height="4"><font size="2" face="Verdana">Email </font>
				<font color="#FF0000" size="2">
				(*)</font></td>
				<td height="4" width="80%" colspan="2">
				<font size="3">
				<input type="text" name="txtEmail" size="52" class="textbox"></font><font color="#FF0000" size="3">&nbsp; 
				</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="15%" height="4"></td>
				<td height="4" width="80%" colspan="2"></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%"><font size="2" face="Verdana">Nội dung </font>
				<font color="#FF0000" size="2">(*)</font></td>
				<td width="64%">
				<font size="3">
				<textarea rows="8" name="txtContent" cols="53" class="textbox"></textarea></font><font color="#FF0000" size="3"> </font></td>
				<td width="16%">
				<font color="#FF0000" size="3">&nbsp;</font></td>
			</tr>
			<tr>
				<td width="5%" height="4"></td>
				<td width="15%" height="4"></td>
				<td width="80%" height="4" colspan="2"></td>
			</tr>
			<tr>
				<td width="99%" colspan="4">
				<p align="center">
				<button name="B1" class="input_button" onclick="JavaScript:CheckInput();">
				<font size="3">&nbsp; Gửi&nbsp;&nbsp;</font>
				</button></font></td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%">&nbsp;</td>
				<td width="80%" colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td width="5%">&nbsp;</td>
				<td width="15%">&nbsp;</td>
				<td width="80%" colspan="2">&nbsp;</td>
			</tr>
		</table>
					</td>
					<input type="hidden" name="category" value="contact">
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