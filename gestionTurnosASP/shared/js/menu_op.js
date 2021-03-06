/*HM_Loader.js
* by Peter Belesis. v4.2.3 020402
* Copyright (c) 2002 Peter Belesis. All Rights Reserved.
*/


if(window.event + "" == "undefined") event = null;
function MenuPopUp(){return false};
function MenuPopDown(){return false};
popUp = MenuPopUp;
popDown = MenuPopDown;

HM_PG_MenuWidth = 150;
HM_PG_FontFamily = "Verdana,Arial,sans-serif";
HM_PG_FontSize = 8;
HM_PG_FontBold = 1;
HM_PG_FontItalic = 0;
HM_PG_FontColor = "blue";
HM_PG_FontColorOver = "blue";
HM_PG_BGColor = "#DDDDDD";
HM_PG_BGColorOver = "#FFCCCC";
HM_PG_ItemPadding = 3;

HM_PG_BorderWidth = 2;
HM_PG_BorderColor = "black";
HM_PG_BorderStyle = "solid";
HM_PG_SeparatorSize = 2;
HM_PG_SeparatorColor = "#d0ff00";

HM_PG_ImageSrc = "HM_More_black_right.gif";
HM_PG_ImageSrcLeft = "HM_More_black_left.gif";
HM_PG_ImageSrcOver = "HM_More_white_right.gif";
HM_PG_ImageSrcLeftOver = "HM_More_white_left.gif";
HM_PG_ImageSize = 5;
HM_PG_ImageHorizSpace = 0;
HM_PG_ImageVertSpace = 2;

HM_PG_KeepHilite = 1; 
HM_PG_ClickStart = 0;
HM_PG_ClickKill = true;
HM_PG_ChildOverlap = 20;
HM_PG_ChildOffset = 10;
HM_PG_ChildPerCentOver = null;
HM_PG_TopSecondsVisible = .5;
HM_PG_StatusDisplayBuild =0;
HM_PG_StatusDisplayLink = 0;
HM_PG_UponDisplay = null;
HM_PG_UponHide = null;
HM_PG_RightToLeft = 0;

HM_PG_CreateTopOnly = 0;
HM_PG_ShowLinkCursor = 1;
HM_PG_NSFontOver = true;

HM_PG_ScrollEnabled = true;
HM_PG_ScrollBarHeight = 14;
HM_PG_ScrollBarColor = "lightgrey";
HM_PG_ScrollImgSrcTop = "HM_More_black_top.gif";
HM_PG_ScrollImgSrcBot = "HM_More_black_bot.gif";
HM_PG_ScrollImgWidth = 9;
HM_PG_ScrollImgHeight = 5;

HM_PG_ScrollBothBars = true;

//HM_a_TreesToBuild = [2]



   HM_DOM = (document.getElementById) ? true : false;
   HM_NS4 = (document.layers) ? true : false;
    HM_IE = (document.all) ? true : false;
   HM_IE4 = HM_IE && !HM_DOM;
   HM_Mac = (navigator.appVersion.indexOf("Mac") != -1);
  HM_IE4M = HM_IE4 && HM_Mac;
 HM_Opera = (navigator.userAgent.indexOf("Opera")!=-1);
 HM_Konqueror = (navigator.userAgent.indexOf("Konqueror")!=-1);

//4.2.3
//HM_IsMenu = !HM_Opera && !HM_Konqueror && !HM_IE4M && (HM_DOM || HM_NS4 || HM_IE4);
HM_IsMenu = !HM_Opera && !HM_IE4M && (HM_DOM || HM_NS4 || HM_IE4 || HM_Konqueror);

HM_BrowserString = HM_NS4 ? "NS4" : HM_DOM ? "DOM" : "IE4";

if(window.event + "" == "undefined") event = null;
function MenuPopUp(){return false};
function MenuPopDown(){return false};
popUp = MenuPopUp;
popDown = MenuPopDown;


HM_GL_MenuWidth          = 150;
HM_GL_FontFamily         = "Arial,sans-serif";
HM_GL_FontSize           = 10;
HM_GL_FontBold           = true;
HM_GL_FontItalic         = false;
HM_GL_FontColor          = "black";
HM_GL_FontColorOver      = "white";
HM_GL_BGColor            = "transparent";
HM_GL_BGColorOver        = "transparent";
HM_GL_ItemPadding        = 3;

HM_GL_BorderWidth        = 2;
HM_GL_BorderColor        = "red";
HM_GL_BorderStyle        = "solid";
HM_GL_SeparatorSize      = 2;
HM_GL_SeparatorColor     = "yellow";

HM_GL_ImageSrc = "HM_More_black_right.gif";
HM_GL_ImageSrcLeft = "HM_More_black_left.gif";

HM_GL_ImageSrcOver = "HM_More_white_right.gif";
HM_GL_ImageSrcLeftOver = "HM_More_white_left.gif";

HM_GL_ImageSize          = 5;
HM_GL_ImageHorizSpace    = 5;
HM_GL_ImageVertSpace     = 5;

HM_GL_KeepHilite         = false;
HM_GL_ClickStart         = false;
HM_GL_ClickKill          = 0;
HM_GL_ChildOverlap       = 40;
HM_GL_ChildOffset        = 10;
HM_GL_ChildPerCentOver   = null;
HM_GL_TopSecondsVisible  = .5;
HM_GL_ChildSecondsVisible = .3;
HM_GL_StatusDisplayBuild = 0;
HM_GL_StatusDisplayLink  = 1;
HM_GL_UponDisplay        = null;
HM_GL_UponHide           = null;

HM_GL_RightToLeft      = false;
HM_GL_CreateTopOnly      = HM_NS4 ? true : false;
HM_GL_ShowLinkCursor     = true;

HM_GL_ScrollEnabled = false;
HM_GL_ScrollBarHeight = 14;
HM_GL_ScrollBarColor = "lightgrey";
HM_GL_ScrollImgSrcTop = "HM_More_black_top.gif";
HM_GL_ScrollImgSrcBot = "HM_More_black_bot.gif";
HM_GL_ScrollImgWidth = 9;
HM_GL_ScrollImgHeight = 5;
HM_GL_ScrollBothBars = false;



// the following function is included to illustrate the improved JS expression handling of
// the left_position and top_position parameters introduced in 4.0.9
// and modified in 4.1.3 to account for IE6 standards-compliance mode
// you may delete if you have no use for it

function HM_f_CenterMenu(topmenuid) {
	var MinimumPixelLeft = 0;
	var TheMenu = HM_DOM ? document.getElementById(topmenuid) : window[topmenuid];
	var TheMenuWidth = HM_DOM ? parseInt(TheMenu.style.width) : HM_IE4 ? TheMenu.style.pixelWidth : TheMenu.clip.width;
//4.2.3
//	var TheWindowWidth = HM_IE ? (HM_DOM ? HM_IEcanvas.clientWidth : document.body.clientWidth) : window.innerWidth;
	var TheWindowWidth = HM_IE ? (HM_DOM ? HM_Canvas.clientWidth : document.body.clientWidth) : window.innerWidth;

	return Math.max(parseInt((TheWindowWidth-TheMenuWidth) / 2),MinimumPixelLeft);
}

if(HM_IsMenu) {
//	document.write("<SCR" + "IPT LANGUAGE='JavaScript1.2' SRC='/ticket/shared/js/HM_Arrays.js' TYPE='text/javascript'><\/SCR" + "IPT>");
	document.write("<SCR" + "IPT LANGUAGE='JavaScript1.2' SRC='/ticket/shared/js/menu_sc.js' TYPE='text/javascript'><\/SCR" + "IPT>");
}

//end