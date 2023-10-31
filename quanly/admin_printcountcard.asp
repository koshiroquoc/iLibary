<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=strSiteName%></title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</head>

<body topmargin="8" leftmargin="8">

<div align="center">

<table border="0" width="797" id="table1" bordercolordark="#808080" cellspacing="0" cellpadding="0" bordercolorlight="#D5F1FF">
	<tr>
		<td>
		<div align="center">
			<table border="0" width="543" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td width ="543" valign="top">
					<div align="center">
						<table border="0" width="543" id="table3" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="3" height="19">
								Phòng GD&amp;ĐT Hải Châu - TP. Đà Nẵng<br>
&nbsp;&nbsp;&nbsp; Trường THCS Trần Hưng Đạo <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								--------</td>
							</tr>
							<tr>
								<td colspan="3" height="19">
								<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								&nbsp;<p style="margin-top: 2px; margin-bottom: 2px" align="center">
								<b><font size="2" color="#FF0000">&nbsp; THỐNG KÊ THẺ THƯ VIỆN NGÀY <%=NgayVN(Now())%></font></b></td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="84">&nbsp;</td>
								<td width="500">
								<table border="1" width="100%" id="table16" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table17" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="10"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center"><b>
												<font size="2" color="#0000FF">THỐNG KÊ 
												TỔNG QUÁT</font></b></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center" style="margin-bottom: 6px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="4">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="23%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="49%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số thẻ 
												thư viện hiện có</font></b></td>
												<td width="3%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="25%"><b>
												<font size="2" color="#0000FF"><%=CountCard()%></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="4">
												<p style="margin-top: 10px; margin-bottom: 10px" align="center">
												<font size="2">
												<img border="0" src="../images/line.gif" width="110" height="5"></font></td>
											</tr>
											<%
												strSQL = "SELECT DISTINCT CLASS_ID FROM CARD"
												Set rsCountCate = Conn.Execute(strSQL)
												Do while Not rsCountCate.Eof
											%>
											<tr>
												<td width="23%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="49%" align="left"><b>
												<font size="2">Lớp:&nbsp;&nbsp;</font><font size="2" color="#003399"><%= rsCountCate("CLASS_ID")%></font></b></td>
												<td width="3%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="25%"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountClassCard(rsCountCate("CLASS_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<%
												rsCountCate.MoveNext
												Loop
												rsCountCate.Close
												Set rsCountCate = Nothing
											%>
											<tr>
												<td width="23%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="49%" align="right">
												&nbsp;</td>
												<td width="3%">&nbsp;</td>
												<td width="25%">&nbsp;</td>
											</tr>
											<tr>
												<td width="101%" align="right" colspan="4">
												&nbsp;</td>
												</tr>
											<tr>
												<td width="23%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="49%" align="right">
												&nbsp;</td>
												<td width="3%">&nbsp;</td>
												<td width="25%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="77">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="84">&nbsp;</td>
								<td width="412">
								<table border="1" width="100%" id="table22" cellspacing="0" cellpadding="0" bordercolorlight="#FFFFFF" bordercolordark="#C0C0C0">
									<tr>
										<td>
										<table border="0" width="100%" id="table23" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="6">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center"><b>
												<font size="2" color="#0000FF">THỐNG KÊ 
												THẺ MƯỢN QUÁ HẠN</font></b></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center" style="margin-bottom: 6px">
												<font size="2">
												<img border="0" src="../images/line.gif" width="175" height="5"></font></td>
											</tr>
											<tr>
												<td colspan="6">
												<p align="center">
												<font size="2">
												<img border="0" src="../images/spacer.gif" width="1" height="4"></font></td>
											</tr>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">
												&nbsp;</td>
												<td width="46%" align="right">
												<p align="left"><b>
												<font size="2">Tổng số thẻ 
												đang mượn quá hạn</font></b></td>
												<td width="2%" align="center">
												<b><font size="2">:</font></b></td>
												<td width="33%" colspan="3"><b>
												<font size="2" color="#0000FF"><%=CountBreaching()%></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<tr>
												<td width="19%" align="right">&nbsp;</td>
												<td width="46%" align="right">
												<p align="left">&nbsp;</td>
												<td width="2%" align="center">&nbsp;</td>
												<td width="12%">&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
											</tr>
											<%
												strSQL = "SELECT DISTINCT CLASS_ID FROM CARD"
												Set rsCategory = Conn.Execute(strSQL)
												Do while Not rsCategory.Eof
												If CountBreachingClass(rsCategory("CLASS_ID"))>0 Then
											%>
											<tr>
												<td width="19%" align="left">
												<p style="margin-left: 20px; margin-top:2px; margin-bottom:2px">&nbsp;</td>
												<td width="46%" align="left"><b>
												<font size="2">Lớp:</font><font size="2" color="#003399">&nbsp;&nbsp;<%=rsCategory("CLASS_ID")%></font></b></td>
												<td width="2%" align="center">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<b><font size="2">:</font></b></td>
												<td width="33%" colspan="3"><b><font color="#0000FF">
												<p style="margin-top: 2px; margin-bottom: 2px">
												<font size="2"><%=CountBreachingClass(rsCategory("CLASS_ID"))%></font></font></b><font size="2">&nbsp;&nbsp;thẻ</font></td>
											</tr>
											<%
												End If
												rsCategory.MoveNext
												Loop
											%>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="46%" align="right">
												&nbsp;</td>
												<td width="2%">&nbsp;</td>
												<td width="12%">&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
											</tr>
											<tr>
												<td width="100%" align="right" colspan="6">
												<p style="margin-right: 3px" align="center">
												</td>
												</tr>
											<tr>
												<td width="19%" align="right">
												<p style="margin-right: 3px">&nbsp;</td>
												<td width="46%" align="right">
												&nbsp;</td>
												<td width="2%">&nbsp;</td>
												<td width="12%">&nbsp;</td>
												<td width="5%">&nbsp;</td>
												<td width="16%">&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
								<td width="77">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="3">
				<table border="0" width="100%" id="table12" cellspacing="0" cellpadding="0">
					<tr>
						<td width="257" bgcolor="#F7F7F7">&nbsp;</td>
						<td width="16" bgcolor="#F7F7F7">
						<font size="2">
						<img border="0" src="../images/print.gif" width="16" height="16"></font></td>
						<td width="123" bgcolor="#F7F7F7">
						<p align="left" style="margin-top: 5px; margin-bottom: 5px; margin-left:8px">
						<b><a href="#" onclick="JavaScript:window.print();">
						<font size="2">In 
						danh sách</font></a></b></td>
						<td width="22" bgcolor="#F7F7F7">
						<p align="center" style="margin-top: 5px; margin-bottom: 5px">
						<font size="2">
						<img border="0" src="../images/close.gif" width="14" height="9"></font></td>
						<td width="62" bgcolor="#F7F7F7">
						<p align="center"><a href="#" onclick="JavaScript:window.close();">
						<b><font size="2">Đóng lại</font></b></a></td>
						<td width="278" bgcolor="#F7F7F7">&nbsp;</td>
					</tr>
				</table>
								</td>
							</tr>
							<tr>
								<td colspan="3">&nbsp;</td>
							</tr>
							</table>
					</div>
					</td>
				</tr>
				</table>
		</div>
		</td>
	</tr>
</table>

</div>

</body>

