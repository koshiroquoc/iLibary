<Script language="JavaScript">
	function openWindow(url) {
	  popupWin = window.open(url,'new_page','width=230,height=140,left=200')
	}
	function openWindow2(url) {
	  popupWin = window.open(url,'new_page','width=400,height=150,left=200')
	}
	function openWindowPrint(url) {
	  popupWin = window.open(url,'new_page','width=600,height=540,left=150,scrollbars=yes,top=0')
	}
	function openWindow3(url) {
	  popupWin = window.open(url,'new_page','width=445,height=235,left=150,top=20')
	}
	function InsertStr(strValue,anh)
	{
		window.opener.document.form1[anh].value=strValue;
	}
	function doSubmit(url)
	{
		document.frmList.action = url;
		document.frmList.submit();
	}
	function cboChange(strComboName){
		document.frmFilter.txtCategoryFilter.value = document.frmList[strComboName].value;
		document.frmFilter.submit();
	}
	function cboChangeClass(strComboName){
		document.frmFilter.txtCategoryFilter.value = document.frmList[strComboName].value;
		document.frmFilter.txtClassID.value = document.frmList[strComboName].value;
		document.frmFilter.submit();
	}
	function handleOver(imgName,srcImg) { 
		if (document.images){ 
			img_on =new Image();
			img_on.src = "../images/" + srcImg + ".gif"
			document[imgName].src=img_on.src;
		}	
	}
	
	function handleOut(imgName,srcImg) {
	 	if (document.images){
		 	img_off=new Image();
		 	img_off.src = "../images/" + srcImg + ".gif"
	 		document[imgName].src=img_off.src;
	 	}	
	}
</Script>