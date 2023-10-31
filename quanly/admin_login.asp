<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_parameter.asp" -->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body bgcolor="#E9E9E9">
<BR>
<br>
<BR>
<div align="center">
<table width="400" border="0" cellspacing="0" bgcolor="#0066CC">
  <TR>
  <TD bgcolor=#E9E9E9 height="40">
  <table border="1" style="border-collapse: collapse" width="100%" id="table1">
    <tr>
      <td>
      <table border="0" style="border-collapse: collapse" width="100%" id="table2">
        <tr>
          <td bgcolor="#0066CC" height="20">
			<p align="center">
			<font face="Verdana" color="#FFFFFF" style="font-size: 9pt">Trường THCS Trưng Vương - Đà 
Nẵng :: Năm học: 2019-2020 </font></td>
        </tr>
        <tr>
          <td>
          <table border="0" style="border-collapse: collapse" width="100%" id="table3" bgcolor="#ECE9D8" cellpadding="0">
          <form name="frmLogin" method="POST" action="admin_process.asp">            
            <tr>
              <td height="66" background="../images/login_top.gif">
				<p align="center"><b><font size="2" color="#FF0000">QUẢN TRỊ HỆ 
				THỐNG THƯ VIỆN</font></b></td>
            </tr>
            <tr>
              <td height="111" background="../images/login_center.gif">
          <table border="0" style="border-collapse: collapse" width="100%" id="table4" cellpadding="0">
            <tr>
              <td colspan="2">
				<p align="center"><b><font size="2" face="Tahoma" color="#FF0000">
'				<%
'					If Request.QueryString("error")<>"" Then
'						strError = Request.QueryString("Error")
'						If strError = 1 Then
'							Response.Write strUsernameBlank
'						Else If strError = 2 Then
'							Response.Write strNoPower
'						Else If strError = 3 Then
'							Response.Write strNoPassword
'						Else If strError = 4 Then
'							Response.Write strNoUserName
'						End If
'				End If
'				%>					
				</font></b>				
			  </td>
            </tr>
            <tr>
              <td align="right" height="6" colspan="2"></td>
            </tr>
            <tr>
              <td width="160" align="right" valign="middle" height="25"><div align="right">
				<span><b>
				<font face="Tahoma" style="font-size: 9pt" color="#333333">Tên 
				truy cập</font><span style="font-size: 9pt"><font color="#333333" face="Tahoma">&nbsp;&nbsp; </font></span>
				</b></span></div></td>
              <td height="25">
                  <input name="txtUsername" type="text" id="user" size="20" style="font-family: Tahoma; font-size: 10pt"></td>
            </tr>
            <tr>
              <td width="160" align="right" valign="middle" height="25">
              <b><font face="Tahoma" style="font-size: 9pt" color="#333333">Mật 
				khẩu</font></b><font face="Tahoma" style="font-size: 9pt" color="#333333"><b>&nbsp;&nbsp; </b></font></td>
              <td height="25">
                  <input name="txtPassword" type="password" id="password" size="20" style="font-family: Tahoma; font-size: 10pt"></td>
            </tr>
            <tr>
              <td colspan="2" align="center" height="35">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Login" type="submit" class="input_button" id="Add" value="Đăng nhập"></td>
            </tr>
          </table>
			  </td>
            </tr>
            <tr>
              <td>
				<img border="0" src="../images/login_end.gif" width="404" height="64"></td>
            </tr>
            </form> 
          </table>
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </TD>
  </TR>
  </table>
</div>
<P style="margin-top: 0; margin-bottom: 0" align="center"> 
&nbsp;</P>
<P style="margin-top: 0; margin-bottom: 0" align="center"> 
<i><font face="Verdana" color="#808080" style="font-size: 9pt"><br>
:: Quản trị Thư viện trực tuyến Trường THCS Trưng Vương - Đà Nẵng ::</font></i></P>
</body>
</html>