ChromelessWindow = function ()
{
	this.URL = null;
	this.title = null;
	this.titleFont = null;
	this.titleSize = 0;
	this.titleColor = null;
	this.titlebarImage = null;
	this.titlebarHeight = 0;
	this.scrollbars = null;
	this.w = 0;
	this.h = 0;
	this.centered = null;
	this.xPos = 0;
	this.yPos = 0;
	this.borderColor = null;
	this.borderSize = 0;
	this.closeBottom = null;
	this.minimizeBottom = null;
	
	this.chromelessCompatible = navigator.appName == "Microsoft Internet Explorer" && parseInt(navigator.appVersion)>=4 ? true:false;
	this.titlebarContent = null;
	this.chromelessWW = null;
	
	
}

ChromelessWindow.prototype.openChromeless = function()
{
	this.runCode();
	
}

ChromelessWindow.prototype.closeChromeless = function()
{
	this.chromelessWW.close();
}

ChromelessWindow.prototype.setURL = function(newURL)
{
	this.URL = newURL;
	if(this.chromelessWW)
	{
		this.chromelessWW.execScript("b.all.canvas.URL='"+newURL+"'");
		this.chromelessWW.execScript("showChromeless()");
	}
}

ChromelessWindow.prototype.moveTo = function()
{
	this.xPos = arguments[0];
	this.yPos = arguments[1];
	if(this.chromelessWW)
	{
		this.chromelessWW.execScript('var cx='+this.xPos+';');
		this.chromelessWW.execScript('var cy='+this.yPos+';');
		this.chromelessWW.execScript("showChromeless()");
	}
}

ChromelessWindow.prototype.resizeTo = function()
{
	this.w = arguments[0];
	this.h = arguments[1];
	this.cH = (this.h-this.titlebarHeight)-this.borderSize;
	this.cW = this.w -(this.borderSize*2);
	if(this.chromelessWW)
	{
		this.chromelessWW.execScript('var cw='+this.w+';');
		this.chromelessWW.execScript('var ch='+this.h+';');
		this.chromelessWW.execScript("b.all.canvas.style.width='"+this.cW+"'");
		this.chromelessWW.execScript("b.all.canvas.style.height='"+this.cH+"'");
		this.chromelessWW.execScript("b.all.closeBottom.style.right='"+this.borderSize+"'");
		this.chromelessWW.execScript("b.all.minimizeBottom.style.right='"+(this.borderSize+20)+"'");
		this.chromelessWW.execScript("showChromeless()");
	}
}



ChromelessWindow.prototype.runCode = function()
{
	if(this.centered=="yes")
	{
		var tempX = this.w;
		var tempY = this.h;
		this.xPos = (screen.availWidth/2)-(tempX/2);
		this.yPos = (screen.availHeight/2)-(tempY/2);
	}
	if (this.chromelessCompatible)
	{
		this.chromelessWW = window.open("about:blank","","top=5000,left=5000,width=100,height=100,scrollbars="+this.scrollbars);
	}
	else
	{
		this.chromelessWW = window.open(this.URL, "chromelessWindow", "top="+this.yPos+",left="+this.xPos+",width="+this.w+",height="+this.h+",toolbar=0,scrollbars="+this.scrollbars+",resizable=0");
	}
	this.cH = (this.h-this.titlebarHeight)-this.borderSize;
	this.cW = this.w -(this.borderSize*2);
	this.c = "<object type='text/html' border='0' style='position:absolute;z-index:2;top:"+this.titlebarHeight+"px;left:"+this.borderSize+"px;width:"+this.cW+"px;height:"+this.cH+"px;' data='"+this.URL+"' id='canvas'></object>";

	if((this.titlebarImage==undefined)||(this.titlebarImage==null)||(this.titlebarImage==""))
	{
		this.titlebarContent="";
		this.tb = false;
	}
	else
	{
		this.titlebarContent="<img id='menu' unselectable='on' src='"+this.titlebarImage+"' style='position:absolute; top:2px; left:5px;z-index:2'>";
		this.tb = true;
	}
	this.menu = "\
		<div id='Title' unselectable='on' style='font-family:Arial;font-size:11px;font-weight:bold;position:absolute;top:5px;left:"+this.borderSize+"px;width:70%;height:15px;cursor:default;font-family:"+this.titleFont+";font-size:"+this.titleSize+";color:"+this.titleColor+";z-index:3'>"+this.title+"</div>\
		<img id='closeBottom' unselectable='on' src='"+this.closeBottom+"' width='11' height='11' style='position:absolute; top:5px; right:"+(this.borderSize)+"px;z-index:3'>\
		"+this.titlebarContent+"\
		<img id='minimizeBottom' unselectable='on' src='"+this.minimizeBottom+"' width='11' height='11' style='position:absolute; top:5px; right:"+(this.borderSize+11)+"px;z-index:3'>\
		<div id='fondo' unselectable='on' style='position:absolute; top:0px; left:0px; width:100%; height:100%; background-color:"+this.borderColor+";z-index:1'></div>";

	if (this.chromelessCompatible)
	{
		this.chromelessWW.moveTo(300,-5000);
		this.chromelessWW.resizeTo(800,600);
		this.chromelessWW.document.writeln('<html>');
		this.chromelessWW.document.writeln('<head>');
		this.chromelessWW.document.writeln('<META HTTP-EQUIV="imagetoolbar" CONTENT="no">');
		this.chromelessWW.document.writeln('<style>');
		this.chromelessWW.document.writeln('BODY{border:0px !important}');
		this.chromelessWW.document.writeln('</style>');
		this.chromelessWW.document.writeln('</head>');
		this.chromelessWW.document.writeln('<body>');
		this.chromelessWW.document.writeln('<script>');
		this.chromelessWW.document.writeln('var c=window.createPopup();');
		this.chromelessWW.document.writeln('var b=c.document.body;');
		this.chromelessWW.document.writeln("var ox=oy=ix=iy=0;");
		this.chromelessWW.document.writeln("var allContent =\""+this.menu+this.c+"\";");
		this.chromelessWW.document.writeln('b.innerHTML=allContent;');
		this.chromelessWW.document.writeln('var cx='+this.xPos+';');
		this.chromelessWW.document.writeln('var cy='+this.yPos+';');
		this.chromelessWW.document.writeln('var cw='+this.w+';');
		this.chromelessWW.document.writeln('var ch='+this.h+';');
		this.chromelessWW.document.writeln('function showChromeless(){');
		this.chromelessWW.document.writeln('c.show(cx,cy,cw,ch);focus()');
		this.chromelessWW.document.writeln('}');
		this.chromelessWW.document.writeln('var mm=false;');
		this.chromelessWW.document.writeln('var cmove=b.all("fondo");');
		this.chromelessWW.document.writeln('var allb=b.all("allborders");');
		this.chromelessWW.document.writeln('var cmove2=(b.all("menu"))?b.all("menu"):"";');
		this.chromelessWW.document.writeln('var ctitle=b.all("Title");');
		this.chromelessWW.document.writeln('ctitle.style.visibility=('+this.tb+')?"hidden":"visible";');
		this.chromelessWW.document.writeln('var cclose=b.all("closeBottom");');
		this.chromelessWW.document.writeln('var cminimize=b.all("minimizeBottom");');
		this.chromelessWW.document.writeln('cclose.onclick=function(){opener=self;self.close()};');
		this.chromelessWW.document.writeln('cminimize.onclick=function(){if(top.opener&&!top.opener.closed){};top.window.blur()};');
		this.chromelessWW.document.writeln('var ox=oy=ix=iy=0;');
		this.chromelessWW.document.writeln('var m=0;');
		this.chromelessWW.document.writeln('ctitle.onmousedown=function(){m=1};');
		this.chromelessWW.document.writeln('ctitle.onmouseup=function(){m=0};');
		this.chromelessWW.document.writeln('cmove.onmousedown=function(){m=1};');
		this.chromelessWW.document.writeln('cmove.onmouseup=function(){m=0};');
		this.chromelessWW.document.writeln('cmove2.onmousedown=function(){m=1};');
		this.chromelessWW.document.writeln('cmove2.onmouseup=function(){m=0};');
		this.chromelessWW.document.writeln('document.body.onclick=function(){ if (m) { focus();cx=ix+event.screenX-ox;cy=iy+event.screenY-oy;c.show(cx,cy,cw,ch) }else { ix=cx; iy=cy; ox=event.screenX; oy=event.screenY } };');
		this.chromelessWW.document.writeln('setInterval("document.body.click()",20);');
		this.chromelessWW.document.writeln('onfocus=function(){showChromeless();}');
		this.chromelessWW.document.writeln('</script>');
		this.chromelessWW.document.writeln('</body>');
		this.chromelessWW.document.writeln('</html>');
		this.chromelessWW.document.location.reload();
		this.chromelessWW.document.title = this.title;
		this.chromelessWW.execScript("focus();");
	}
	else
	{
		this.chromelessWW.document.title = this.title;
	}
}
