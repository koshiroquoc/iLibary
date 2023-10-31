<%
	strSiteName = "H&#7879; th&#7889;ng h&#7895; tr&#7907; th&#244;ng tin."
	strNoValue	= "B&#7841;n ch&#432;a nh&#7853;p &#273;&#7847;y &#273;&#7911; d&#7919; li&#7879;u."
	strExistValue	= "M&#227; d&#7919; li&#7879;u &#273;&#227; t&#7891;n t&#7841;i."
	strNotEqual	= "M&#7853;t kh&#7849;u nh&#7853;p l&#7841;i kh&#244;ng kh&#7899;p."
	strExitsUser= "Ng&#432;&#7901;i d&#249;ng ho&#7863;c nh&#243;m &#273;&#227; &#273;&#432;&#7907;c ph&#226;n quy&#7873;n."
	strUsernameBlank= "B&#7841;n ch&#432;a nh&#7853;p T&#234;n truy c&#7853;p."
	strNoUsername = "T&#234;n truy c&#7853;p kh&#244;ng t&#7891;n t&#7841;i."
	strNoPassword = "B&#7841;n nh&#7853;p sai m&#7853;t kh&#7849;u."
	strNoPower = "B&#7841;n kh&#244;ng c&#243; quy&#7873;n qu&#7843;n l&#253;."
	strNotPower = "B&#7841;n kh&#244;ng c&#243; quy&#7873;n qu&#7843;n l&#253;."
	strKeyBlank = "B&#7841;n ch&#432;a nh&#7853;p kh&#243;a t&#236;m ki&#7871;m ho&#7863;c ch&#7885;n c&#225;c t&#249;y ch&#7885;n."
	strExitCard = "Mã th&#7867; này không t&#7891;n t&#7841;i! Xin vui lòng xem l&#7841;i!"
	strExitBook = "Mã sách này không t&#7891;n t&#7841;i ho&#7863;c không có m&#432;&#7907;n sách! Xin vui lòng xem l&#7841;i!"
	strYearValid = "N&#259;m không th&#7875; b&#7857;ng n&#259;m hi&#7879;n t&#7841;i!"
	strEndBook = "Sách có nh&#432;ng &#273;ã &#273;&#432;&#7907;c m&#432;&#7907;n h&#7871;t! Xin ch&#7901; khi sinh viên khác tr&#7843; l&#7841;i!"
	strDoneBook = "Th&#7867; &#273;ã m&#432;&#7907;n sách! Hãy tr&#7843; sách tr&#432;&#7899;c khi m&#432;&#7907;n ti&#7871;p!"
	strNoBorrow = "Th&#7867; này không có m&#432;&#7907;n sách! Không th&#7875; th&#7921;c hi&#7879;n vi&#7879;c tr&#7843; sách!"
	strExisClass = "M&#7895;i l&#7899;p ch&#7881; import 1 l&#7847;n, n&#7871;u thi&#7871;u b&#7841;n c&#7847;n b&#7893; sung &#273;&#7875; tránh trùng l&#7863;p"
	strInvalid = "N&#7871;u nhân viên thì 2 ký t&#7921;, h&#7885;c sinh thì 4 ký t&#7921;, sinh viên thì 6 ký t&#7921;!"
	Function GetTitle(strCatalog)
		If strCatalog ="" Then
			strCatalog ="default"
		End If	
		If strCatalog = "default" Then
			strTitle = "Trang ch&#7911;"
		End If	
		If strCatalog = "news" Then
			strTitle = "Tin t&#7913;c"
		End If	
		If strCatalog = "search" Then
			strTitle = "T&#236;m ki&#7871;m s&#225;ch"
		End If
		If strCatalog = "bookonline" Then
			strTitle = "S&#225;ch tr&#7921;c tuy&#7871;n"
		End If		
		If strCatalog = "soft" Then
			strTitle = "Ti&#7879;n &#237;ch"
		End If		
		If strCatalog = "relax" Then
			strTitle = "G&#243;c th&#432; gi&#7843;n"
		End If		
		If strCatalog = "inform" Then
			strTitle = "Th&#244;ng b&#225;o"
		End If		
		If strCatalog = "category" Then
			strTitle = "Danh m&#7909;c s&#225;ch"
		End If		
		If strCatalog = "schedule" Then
			strTitle = "L&#7883;ch m&#7903; c&#7917;a"
		End If	
		If strCatalog = "help" Then
			strTitle = "Tr&#7907; gi&#250;p"
		End If
		If strCatalog = "bookresult" Then
			strTitle = "K&#7871;t qu&#7843; t&#236;m ki&#7871;m"
		End If	
		If strCatalog = "searchbook" Then
			strTitle = "T&#236;m ki&#7871;m s&#225;ch"
		End If
		If strCatalog = "error" Then
			strTitle = "L&#7895;i thao t&#225;c"
		End If
		If strCatalog = "listdoc" Then
			strTitle = "Danh s&#225;ch t&#224;i li&#7879;u"
		End If			
		If strCatalog = "docdetail" Then
			strTitle = "N&#7897;i dung t&#224;i li&#7879;u"
		End If
		If strCatalog = "listcatenews" Then
			strTitle = "Danh s&#225;ch nh&#243;m tin"
		End If	
		If strCatalog = "newsdetail" Then
			strTitle = "N&#7897;i dung tin t&#7913;c"
		End If		
		If strCatalog = "listsoft" Then
			strTitle = "Ti&#7879;n &#237;ch h&#7895; tr&#7907;"
		End If	
		If strCatalog = "notice" Then
			strTitle = "Thông báo"
		End If
		If strCatalog = "notidetail" Then
			strTitle = "Chi ti&#7871;t thông báo"
		End If
		If strCatalog = "listcatesoft" Then
			strTitle = "Dowload ti&#7879;n ích"
		End If
		If strCatalog = "catebook" Then
			strTitle = "Danh m&#7909;c sách"
		End If
		If strCatalog = "listrelax" Then
			strTitle = "Góc gi&#7843;i trí"
		End If						
		If strCatalog = "relaxdetail" Then
			strTitle = "N&#7897;i dung gi&#7843;i tr&#237;"
		End If								
		If strCatalog = "listcatebook" Then
			strTitle = "Danh m&#7909;c sách"
		End If				
		If strCatalog = "searchhome" Then
			strTitle = "K&#7871;t qu&#7843; tìm ki&#7871;m"
		End If	
		If strCatalog = "resultsoft" Then
			strTitle = "K&#7871;t qu&#7843; tìm ki&#7871;m"
		End If														
		If strCatalog = "resultdoc" Then
			strTitle = "K&#7871;t qu&#7843; tìm ki&#7871;m"
		End If	
		If strCatalog = "contact" Then
			strTitle = "Liên h&#7879;"
		End If			
		If strCatalog = "introduce" Then
			strTitle = "Gi&#417;&#769;i thiê&#803;u"
		End If	
		If strCatalog = "diagram" Then
			strTitle = "S&#417; &#273;ô&#768; website"
		End If							
		If strCatalog = "rule" Then
			strTitle = "N&#7897;i quy th&#432; vi&#7879;n"
		End If	
		If strCatalog = "catedoc" Then
			strTitle = "Danh m&#7909;c tài li&#7879;u"
		End If																			
		If strCatalog = "listcatedoc" Then
			strTitle = "Chuyên m&#7909;c tài li&#7879;u"
		End If
		If strCatalog = "charbook" Or strCatalog = "advancebook" Then
			strTitle = "K&#7871;t qu&#7843; tìm ki&#7871;m"
		End If
		If strCatalog = "advancesearch" Then
			strTitle = "Tìm ki&#7871;m nâng cao"
		End If	
		If strCatalog = "resultnotice" Then
			strTitle = "Thông báo"
		End If
		If strCatalog = "homebook" Then
			strTitle = "Th&#244;ng tin s&#225;ch"
		End If	
		If strCatalog = "link" Then
			strTitle = "Liên k&#7871;t"
		End If	
		GetTitle =strTitle	
	End Function
%>