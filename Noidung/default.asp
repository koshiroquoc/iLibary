<%Session.Codepage=65001%>

<!--#INCLUDE FILE="../include/inc_function.asp" -->
<!--#INCLUDE FILE="public_check.asp" -->
<!--#INCLUDE FILE="../include/inc_parameter.asp" -->
<html>

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>HỆ THỐNG QUẢN LÝ THƯ VIỆN TRỰC TUYẾN - TRƯỜNG THCS TRƯNG VƯƠNG - ĐÀ NẴNG :: LIBRARY ONLINE :: </title>
</head>
<% 
	strCatalog=Request.Querystring("name")
	strTitle = GetTitle(strCatalog) 
%>
<link rel="stylesheet" type="text/css" href="../css/public.css">
<body bgcolor="#E7EEDF" background="../images/nen-xam.gif" >
<div align="center">
<table border="0" width="984" id="table1" cellspacing="0" cellpadding="0" >
	<tr>
		<td>
			<table border="0" width="984" id="table2" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td colspan="5" bgcolor="#F2F9FF">
					<p align="center"><!--#INCLUDE FILE="public_header.asp" --></td>
				</tr>
				<tr>
				
					<td width="162" valign="top" bgcolor="#F2F9FF">
						<!--#INCLUDE FILE="public_menu.asp" -->
						<!--#INCLUDE FILE="public_search.asp" -->
					</td>
					<td width="13" bgcolor="#FFFFFF"></td>
					<td width="667" valign="top" bgcolor="#FFFFFF">
					<table border="0" width="100%" id="table3" bordercolorlight="#999999" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
						<tr>
							<td height="26">
							<table border="0" width="106%" id="table4" cellspacing="0" cellpadding="0" height="16">
								<tr>
									<td width="19">
									<p align="center">
									<font color="#FF0000">
									<img border="0" src="../images/mod_titl.gif" width="8" height="9"></font></td>
									<td width ="246"><b><span style="font-size: 10pt">
									<font color="#FF0000"><%=strTitle%></font></span></b></td>
									<td width="363">
									<p align="right"><font color="#0000FF"><span style="font-size: 10pt">Hôm nay, 
									ngày <%=NgayVN_Text(Now())%></span></font></td>
									<td width="67">&nbsp;</td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td>
							<p style="margin: 4px">
								<%	If strCatalog = "default" Then	%>
									<!--#INCLUDE FILE="public_default.asp" -->
								<%	End If	%>
								<%	If strCatalog = "news" Then	%>
									<!--#INCLUDE FILE="public_listnews.asp" -->
								<%	End If	%>	
								<% If strCatalog="newsdetail" Then%>
									<!--#INCLUDE FILE="public_newsdetail.asp" -->
								<% End If %>				
								<% If strCatalog="listcatenews" Then%>
									<!--#INCLUDE FILE="public_listcatenews.asp" -->
								<% End If %>
								<% If strCatalog="searchbook" Then%>
									<!--#INCLUDE FILE="public_searchbook.asp" -->
								<% End If %>
								<% If strCatalog="bookresult" Then%>
									<!--#INCLUDE FILE="public_resultbook.asp" -->
								<% End If %>
								<% If strCatalog="error" Then%>
									<!--#INCLUDE FILE="public_error.asp" -->
								<% End If %>							
								<% If strCatalog="listdoc" Then%>
									<!--#INCLUDE FILE="public_listdoc.asp" -->
								<% End If %>									
								<% If strCatalog="listcatedoc" Then%>
									<!--#INCLUDE FILE="public_listcatedoc.asp" -->
								<% End If %>																	
								<% If strCatalog="docdetail" Then%>
									<!--#INCLUDE FILE="public_docdetail.asp" -->
								<% End If %>																	
								<% If strCatalog="listsoft" Then%>
									<!--#INCLUDE FILE="public_listsoft.asp" -->
								<% End If %>
								<% If strCatalog="listcatesoft" Then%>
									<!--#INCLUDE FILE="public_listcatesoft.asp" -->
								<% End If %>	
								<% If strCatalog="listrelax" Then%>
									<!--#INCLUDE FILE="public_listrelax.asp" -->
								<% End If %>
								<% If strCatalog="catebook" Then%>
									<!--#INCLUDE FILE="public_catebook.asp" -->
								<% End If %>																																										
								<% If strCatalog="notice" Then%>
									<!--#INCLUDE FILE="public_notice.asp" -->
								<% End If %>																																										
								<% If strCatalog="notidetail" Then%>
									<!--#INCLUDE FILE="public_notidetail.asp" -->
								<% End If %>																																										
								<% If strCatalog="listcatebook" Then%>
									<!--#INCLUDE FILE="public_listcatebook.asp" -->
								<% End If %>																																										
								<% If strCatalog="relaxdetail" Then%>
									<!--#INCLUDE FILE="public_relaxdetail.asp" -->
								<% End If %>																																										
								<% If strCatalog="schedule" Then%>
									<!--#INCLUDE FILE="public_schedule.asp" -->
								<% End If %>
								<% If strCatalog="resultsoft" Then%>
									<!--#INCLUDE FILE="public_resultsoft.asp" -->
								<% End If %>																																										
								<% If strCatalog="searchhome" Then%>
									<!--#INCLUDE FILE="public_searchhome.asp" -->
								<% End If %>																																										
								<% If strCatalog="resultdoc" Then%>
									<!--#INCLUDE FILE="public_resultdoc.asp" -->
								<% End If %>																																										
								<% If strCatalog="contact" Then%>
									<!--#INCLUDE FILE="public_contact.asp" -->
								<% End If %>																																										
								<% If strCatalog="introduce" Then%>
									<!--#INCLUDE FILE="public_introduce.asp" -->
								<% End If %>		
								<% If strCatalog="diagram" Then%>
									<!--#INCLUDE FILE="public_diagram.asp" -->
								<% End If %>																																										
								<% If strCatalog="rule" Then%>
									<!--#INCLUDE FILE="public_rule.asp" -->
								<% End If %>																																										
								<% If strCatalog="catedoc" Then%>
									<!--#INCLUDE FILE="public_catedoc.asp" -->
								<% End If %>
								<% If strCatalog="charbook" Then%>
									<!--#INCLUDE FILE="public_charbook.asp" -->
								<% End If %>
								<% If strCatalog="advancesearch" Then%>
									<!--#INCLUDE FILE="public_advancesearch.asp" -->
								<% End If %>
								<% If strCatalog="advancebook" Then%>
									<!--#INCLUDE FILE="public_advancebook.asp" -->
								<% End If %>
								<% If strCatalog="resultnotice" Then%>
									<!--#INCLUDE FILE="public_resultnotice.asp" -->
								<% End If %>
								<% If strCatalog="homebook" Then%>
									<!--#INCLUDE FILE="public_homebook.asp" -->
								<% End If %>
								<% If strCatalog="link" Then%>
									<!--#INCLUDE FILE="public_link.asp" -->
								<% End If %>								
							</td>
						</tr>
					</table>
					</td>
					<td width="5" bgcolor="#F2F9FF"></td>
					<td width="155" valign="top" bgcolor="#F2F9FF">
						<!--#INCLUDE FILE="public_inform.asp" -->
						
					</td>
				</tr>
				<tr>
					<td colspan="5" bgcolor="#F2F9FF"><!--#INCLUDE FILE="public_footer.asp" --></td>
				</tr>
			</table>
		</td>
	</tr>
</table>

</div>

</body>

</html>
<noscript><object><layer><style><title><xml><apple t>
<noembed><noframes></noscript>