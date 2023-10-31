
<head>
<meta http-equiv="Content-Language" content="en-us">
</head>

<table border=0 cellpadding=0 style="width:280px; border:1px Double darkgray" bgcolor="buttonface">
	<style>
		.command{width:100;background-color:black;color:#00ff00}
		textarea	{border-collapse: collapse; border-style: inset; border-width: 1px; background-color:'#F3F3F3'}
		.select		{font:10px verdana;color=blue;background-color:white;border:1px solid darkblue;}
		select		{border-collapse: collapse; border-style: inset; border-width: 1px}		
		.hdr	{font:18px verdana;color=red;}
		.bld	{font:13px verdana;color=darkviolet;font-weight:bold;}
		.error	{color:red;display:"none"}
		button		{font:10px verdana;}
		button.btn	{border-width:1;width:26px;height:24px; }
		button.btndn	{border-width:1;width:26px;height:24px;border-style:inset;background-color:buttonhighlight; }
		button.btnna	{border-width:1;width:26px;height:24px;filter:alpha(opacity=25);}
		div.Drop	{position:absolute;border:2px solid darkblue;background-color:white;filter:alpha(opacity=100,finishopacity=60,style=1);padding-left:3px;padding-right:3px;padding-top:3px;padding-bottom:3px;}
		img.out		{border:2px solid white;}
		.over	{cursor:hand;border:2px solid blue;background-color:black}
		a.menue		{position:"relative";color:blue;font-size:12px;text-decoration:none;padding-left:3px;padding-right:3px;padding-top:1px;padding-bottom:1px;}
		a.menue:hover		{color:white;background-color:blue;filter:alpha(opacity=0,finishopacity=100,style=1);}
	</style>
<tr><td width="100%">
	<button title="Bold" id="cmdBold" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_bold.gif" width="23" height="22"></button>
	<button title="Italic" id="cmdItalic" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_italic.gif"></button>
	<button title="Underline" id="cmdUnderline" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_underline.gif"></button>
	<button title="StrikeThrough" id="cmdStrikeThrough" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_strike.gif"></button>
	<button title="Justify Left" id="cmdJustifyLeft" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_align_left.gif"></button>
	<button title="Justify Center" id="cmdJustifyCenter" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_align_center.gif"></button>
	<button title="Justify Right" id="cmdJustifyRight" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_align_right.gif"></button>
	<button title="Ordered List" id="cmdInsertOrderedList" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_list_num.gif"></button>
	<button title="Bulleted List" id="cmdInsertUnorderedList" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_list_bullet.gif"></button>
	<button title="Decrease Indent" id="cmdOutdent" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_indent_less.gif"></button>
	<button title="Increase Indent" id="cmdIndent" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_indent_more.gif"></button>
</td></tr>

<tr><td width="100%">
	<button title="Cut" id="cmdCut" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_cut.gif"></button>
	<button title="Copy" id="cmdCopy" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_copy.gif"></button>
	<button title="Paste" id="cmdPaste" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_paste.gif"></button>
	<button title="Font Name" id="cmdFontName" class="btn" onClick="turnDiv('Name')" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_font_face.gif"></button>
	<button title="Font Size" id="cmdFontSize" class="btn" onClick="turnDiv('Size')" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_font_size.gif"></button>
	<button title="Paragraph Style" id="cmdParagraphStyle" class="btn" onClick="turnDiv('Paragraph')" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_font_paragraph.gif"></button>
	<button title="Font Color" id="cmdForeColor" class="btn" onClick="setColor('ForeColor')" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_color_fg.gif"></button>
	<button title="Background Color" id="cmdBackColor" class="btn" onClick="setColor('BackColor')" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_color_bg.gif"></button>
	<button title="Undo" id="cmdUndo" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_undo.gif"></button>
	<button title="Redo" id="cmdRedo" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_redo.gif"></button>
	<button title="Preview" id="cmdAbout" class="btn" onClick="previewHTML()" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_preview.gif"></button>
</td></tr>

<tr><td width="100%">
	<button title="Print" id="cmdPrint" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_print.gif"></button>
	<button title="Insert Marquee: Double Click into it after insert and ENTER text." id="cmdInsertMarquee" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_marquee.gif"></button>
	<button title="Horizontal Rule" id="cmdInsertHorizontalRule" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_hr.gif"></button>
	<button title="Insert WEB'Link" id="cmdCreateLink" class="btn" onClick="createLink()" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_link.gif"></button>
	<button title="Subscript" id="cmdSubScript" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_sub.gif"></button>
	<button title="Superscript" id="cmdSuperScript" class="btn" onClick="goAction(this.id)" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_format_sup.gif"></button>
	<button title="Swapmode" id="cmdSuperScript" class="btn" onClick="swapmode()" unselectable="on">
	<img onmouseout='this.style.background=""' onmouseover='this.style.background="#FFFFFF"' src="../images/editor/ed_swapmode.gif"></button>

	</td>
</tr>

<tr>
	<td width="100%">
		<table border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td style="width:79px"></td>
			<td style="width:23px">
				<div class="Drop" id="divFName" style="display:none">
				<table border=0 cellspacing=0 cellpadding=0>
				<%
				Dim arrFonts, nInt
				arrFonts= Array("Arial", "Arial Black", "Arial Narrow", "Comic Sans MS", "Courier New", "Times New Roman", "Tahoma", "Verdana", "webdings", "Wingdings")
				For nInt=0 To 9
					Response.Write "<tr><td align='center'><a class='menue' style='font:12px " & arrFonts(nInt) & ";width:120px;' href='javascript:setFName(FN" & nInt & ".innerText)'><font id=FN" & nInt & ">" & arrFonts(nInt) & "</font></a></td></tr>"
				Next
				%>
				</table>
				</div>
			</td>
			<td style="width:24px">
				<div class="Drop" id="divFSize" style="display:none">
				<table border=0 cellspacing=0 cellpadding=0>
				<%
				For nInt=1 To 7
					Response.Write "<tr><td align='center'><a class='menue' style='width:150px;' href='javascript:setFSize(" & nInt & ")'><font size=" & nInt & ">size " & nInt & "</font></a></td></tr>"
				Next
				%>
				</table>
				</div>
			</td>
			<td style="width:82px">
				<div class="Drop" id="divFParagraph" style="display:none">
				<table border=0 cellspacing=0 cellpadding=0>
				<%
				For nInt=6 To 1 Step -1
					Response.Write "<tr><td align='center'><a class='menue' style='width:230px;' href='javascript:setP(""Heading " & nInt & """)'><h" & nInt & ">&#272;&#7883;nh d&#7841;ng " & nInt & "</h></a></td></tr>"
				Next
				%>
				</table>
				</div>
			</td>
		</tr>
	</table>
</td></tr>

<tr><td width="100%" bgcolor="buttonface">
</td></tr>

<tr id="ifrm" style="display:inline;"><td width="100%" bgcolor="buttonface"><IFRAME name="rTE" ID="rTE" width="100%" height="178"></IFRAME></td></tr>

<script language=javaScript>
	function swapmode()
	{
		if (getObj("ifrm").style.display=="inline")	{
			getObj("ifrm").style.display="none";
			getObj("txtarea").style.display="inline";
		}
		else{
			getObj("ifrm").style.display="inline";
			getObj("txtarea").style.display="none";
		}
	}	

	function turnDivOFF(aID)
	{
		getObj("divF"+ aID).style.display="none";
	}
	function turnDiv(aID)
	{
		var arrID= ["Name", "Size", "Paragraph"]
		var arrHeight= [275, 300, 250];
		var aHeight;
		for (i=0;i<arrID.length;i++)
		{
			if (aID!=arrID[i]) turnDivOFF(arrID[i])
			else aHeight= arrHeight[i];
		}
		if (aID==null) return;
		var objDiv= getObj("divF"+ aID);
		if (objDiv.style.display=="inline") {turnDivOFF(aID); return;}
		objDiv.style.left= event.screenX;
		objDiv.style.top= event.screenY-100;
		objDiv.style.display="inline";
	}
	function DoCom(cCmd, cArg)
	{
		rTE.focus();
		rTE.document.execCommand(cCmd, false, cArg);
		var aHeight=null;
		turnDiv(null);
		updateToolBar();
	}
	function GetCom(cCmd) {rTE.focus();return rTE.document.queryCommandValue(cCmd);}
	function goAction(oID) {DoCom(oID.substr(3,oID.length-2),null);}
	function setFName(cVal) {DoCom("FontName", cVal);}
	function setFSize(cVal) {DoCom("FontSize", cVal);}
	function setP(cVal){DoCom("formatBlock", cVal);}
	function createLink() {rTE.focus();rTE.document.execCommand('CreateLink',1);}
	function Dec2RGB(value)
	{
		var hex_string= "";
		for (var hexpair= 0;hexpair<3;hexpair++)
		{
			var byte= value & 0xFF;
			value>>= 8;
			var nybble2= byte & 0x0F;
			var nybble1= (byte>> 4) & 0x0F;
			hex_string+= nybble1.toString(16);
			hex_string+= nybble2.toString(16);
		}
		return hex_string.toUpperCase();
	}
	function setColor(cKind)
	{
		turnDiv(null);
		var posX= event.screenX- 10;
		var posY= event.screenY+ 10;
		var wPosition= "dialogLeft:"+posX+";dialogTop:"+posY;
		var oldcolor= Dec2RGB(GetCom(cKind));
		var newcolor= showModalDialog("Images/editor/color.htm", oldcolor,
						"dialogLeft:"+posX+";dialogTop:"+ posY
						+ ";dialogWidth:238px; dialogHeight: 187px; "
						+ "resizable: no; help: no; status: no; scroll: no; ");
		if (newcolor != null) {DoCom(cKind, newcolor);}
	}
	function updateToolBar()
	{
		turnDiv(null);
		document.all["txtMessage"].value= rTE.document.body.innerHTML;
		var arrObjs= ["Bold", "Italic", "Underline", "StrikeThrough", "SubScript", "SuperScript",
				"JustifyLeft", "JustifyCenter", "JustifyRight", "InsertOrderedList", "InsertUnorderedList"]
		for (i=0;i<arrObjs.length;i++)
		{
			if (!GetCom(arrObjs[i])) document.all["cmd"+ arrObjs[i]].className= "btn"
			else document.all["cmd"+ arrObjs[i]].className= "btnDN";
		}
	}
	function strupdate()
	{
		rTE.document.body.innerHTML = document.all["txtMessage"].value;
	}

	function strupdate2()
	{
		document.all["txtMessage"].value = rTE.document.body.innerHTML;
	}
	
	var BodyTag= "<BODY MultiSPACE STYLE='OverFlow:Auto;border-collapse: collapse; border-style: inset; border-width: 1px;Margin-Left:2px; Margin-Top:2px; Margin-Right:2px; Margin-Bottom:2px;'><P ALIGN=Justify STYLE='Font:12px Arial;margin-top: 2px; margin-bottom: 2px'>"
	rTE.document.open();
	rTE.document.write(BodyTag);
	rTE.document.close();
	rTE.document.designMode="On";
	rTE.document.onkeyup= function() {updateToolBar();}
	rTE.document.onmouseup= function() {updateToolBar();}
	rTE.document.onmouseout= function() {strupdate2();}	

	function turnTD(tdName, lDisplay)
	{
		var objTD= getObj(tdName)
		if (lDisplay) objTD.style.display= "inline"
		else objTD.style.display= "none";
		return lDisplay;
	}
	function getObj(objName){return document.getElementById(objName);}
	function countChar(cStr, cChr)
	{
		var nCnt=0;
		for (i=0;i<=cStr.length;i++) if (cStr.charAt(i)==cChr) nCnt++;
		return nCnt
	}
	function changeAmount(nA)
	{
		if (nA=="") nA=0;
		var sNum= "0123456789";
		var sHTML1= "<input type=file name='fAttachment"
		var sHTML2= "' style='width:392px'>";
		if (sNum.search(nA+ "")==-1) nA=0;
		var objDiv= getObj("divFAttachment")
		var sRes=""
		for (i=1;i<=nA;i++)
		{
			if (i>1) sRes+= "<br>";
			sRes+= sHTML1+ i+ sHTML2;
		}
		document.getElementById("divFAttachment").innerHTML=sRes;
		return nA;
	}
	function previewHTML()
	{
		var popup = open('', 'ColorPicker',
                  "location=no,menubar=no,toolbar=no,directories=no,status=no," +
                  "top=200,left=200,height=200,width=400,resizable=no,scrollbars=yes");
		popup.document.write("<heaad><title>Xem Noi Dung Truoc Khi Gui </title></head><body>" + rTE.document.body.innerHTML);
	}
	document.oncontextmenu=function() {return false};
</script>
<tr style="display:none;" id="txtarea"><td>
<textarea style="width:285px; height:180px" onkeyup="strupdate()" name="txtContent" id="txtMessage"></textarea>
</td></tr>
</table>