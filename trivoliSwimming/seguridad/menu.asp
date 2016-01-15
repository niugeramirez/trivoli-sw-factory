<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!------------------------------------------------------------------------------
Archivo       : menu.asp
Descripcion   : Armado del menu
Creacion      : 
Autor         : 
Modificacion  :
------------------------------------------------------------------------------->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
Set rs4 = Server.CreateObject("ADODB.RecordSet")
Set mirs = Server.CreateObject("ADODB.RecordSet")

Dim l_rs
Dim l_sql
Dim l_username
Dim l_planta
Dim mirs

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT usrnombre from user_per where upper(iduser) = '" & uCase(Session("Username")) & "'"
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0
l_username = l_rs(0)
if l_rs.eof then 
 	response.write "<script>window.alert('Usuario no autorizado');"
 	response.write "document.location='blanc.html';</script>"
end if
l_rs.Close
l_rs = nothing

%>
<html>
<head>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
function about(){
	ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');
	//alert('RHPro ® - Supervisor v.3.11\nCopyright Heidt & Asociados S.A.\nNoviembre 2002. Todos los derechos reservados.');
}

function Cerrar(){
	if (confirm('¿ Desea salir de Supervisor ?') == true){
				//parent.main.location = "closesession.asp";
	}
}
</script>
<style type="text/css">
	.cswmDock{display:none;position:absolute;top:0px;left:0px;border:dotted #000000 1px;padding:0px;width:100%;}
	.cswmButtons{position:absolute;z-index:999;top:0px;left:0px;background-color:#d4d0c8;border-top: solid #d4d0c8 1px;border-left: solid #d4d0c8 1px;border-bottom: solid #000000 1px;border-right: solid #000000 1px;padding:0px;cursor:default;width:auto;}
	.cswmInnerBorder{background-color:#d4d0c8;border-top: solid #ffffff 1px;border-left: solid #ffffff 1px;border-bottom: solid #808080 1px;border-right: solid #808080 1px;padding:0px;width:100%;}
	.cswmButton{background-color:#d4d0c8;border-top: solid #d4d0c8 1px;border-left: solid #d4d0c8 1px;border-bottom: solid #d4d0c8 1px;border-right: solid #d4d0c8 1px;color:#000000;font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif;font-size:11px;font-style:normal;text-decoration:none;font-weight:normal;text-align:center;padding-top:4px;padding-bottom:4px;padding-left:6px;padding-right:6px;}
	.cswmHandle{background-color:#d4d0c8;border-top: solid #ffffff 1px;border-left: solid #ffffff 1px;border-bottom: solid #808080 1px;border-right: solid #808080 1px;cursor:move;width:3px;}
	.cswmNNDck{position:absolute;border:outset #808080 1px;width:100%;color:#d0d0d0;font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif;font-size:11px;font-style:normal;text-decoration:none;font-weight:normal;text-align:center;text-decoration:none;padding:3px;}
	.cswmNNBtns{position:absolute;}
	.cswmNNBtn{border-style:outset;border-color:#d4d0c8;border-width:1px;}
	.cswmNNBtnTxt{color:#000000;font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif;font-size:11px;font-style:normal;text-decoration:none;font-weight:normal;text-align:center;padding:3px;}
	.cswmNNHnd{color:#000000;font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif;font-size:11px;font-style:normal;text-decoration:none;font-weight:normal;text-align:center;padding:3px;}
	.cswmItem {font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif; font-size:11px; font-weight:normal; font-style:normal; color:#000000; text-decoration:none; padding:3 10 3 4}
	.cswmItemOn {font-family:"MS Sans Serif", Arial, Helvetica, Tahoma, sans-serif; font-size:11px; font-weight:normal; font-style:normal; color:#000000; text-decoration:none; padding:3 10 3 4}
	.cswmExpand {cursor:default}
	.cswmPopupBox {cursor:default; position:absolute; left:-500; display:none; z-index:1999}
	.cswmDisabled {color:#a5a6a6}
</style>
<LINK href="/trivoliSwimming/shared/css/ie4.css" rel=stylesheet>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script language="javascript" type="text/javascript">

<!--
var cswmDetectedBrowser = 'IE6DHTML';
var cswmMBZ=false;
var cswmCSDS=false;
var cswmOM="document.all.";
var BgCo=".style.backgroundColor";
var cswmCo=".style.color";
var cswmDi=".style.display";
var cswmTI="";
var cswmClkd=-1;
var cswmPI=new Array();
var cswmPx=new Array();
var cswmPy=new Array();
var cswmNH=new Array();
var cswmPW=0;
var cswmPH=0;
var cswmSPnt="";
var cswmDir="";
var cswmMB=0;
var cswmSI="";
var cswmSE=new Object();
var cswmSEL=0;
var cswmSET=0;
var cswmSEH=0;
var cswmSEW=0;
var cswmBW=0;
var cswmBH=0;
var cswmAR=0;
var cswmAB=0;
var cswmSLA=0;
var cswmSTA=0;
var cswmExIS=new Image();
cswmExIS.src="/trivoliSwimming/shared/images/Popup.gif";
var cswmExdIS=new Image();
cswmExdIS.src="/trivoliSwimming/shared/images/Popup.gif";
var cswmCTH=true;
var cswmXOff=0;
var cswmYOff=0;
var cswmFP=0;
var cswmSH=false;
var cswmSTI=0;
var cswmSdw = new Array();

function cswmT(ms){
	if(ms!="off"){
		if(cswmCTH!=0){
			cswmTI=setTimeout("cswmHP(0);",ms);
		}
	}else{
		clearTimeout(cswmTI);
	}
}

function cswmST(l,g,i){
	if(i){
		cswmSTI = setTimeout("cswmHP("+l+");cswmSP("+g+",'"+i+"');",350);
	}else{
		if(l){
			cswmSTI = setTimeout("cswmHP("+l+");",350);
		}else{
			clearTimeout(cswmSTI);
		}
	}
}

function cswmShow(id,srcid,relpos,offsetX,offsetY,fixedpos){
	clearTimeout(cswmTI);
	if(cswmClkd!=id){
		cswmHP(0);
		cswmSI=srcid;
		cswmSPnt=relpos;
		cswmClkd=id;
		cswmDir="right";
	    if(document.all["cswmPopup"+id]){
			if(offsetX)
				cswmXOff=offsetX;
			if(offsetY)
				cswmYOff=offsetY;
			if(fixedpos)
				cswmFP=fixedpos;
      		cswmButtonClickState=true;
      		cswmSP(id);
		}
	}
}

function cswmHide(){
	cswmTI=setTimeout("cswmHP(0);", 350);
}

function cswmHiI(id,l){
  var d;
  d=document;
  if(!d.all["cswmItem"+id])
  {
    return;
  }
  var bgco;
  try
  {
    bgco =  d.all['cswmItem'+id].getAttribute('cswmSelColor');
  }
  catch(e)
  {
    bgco = false;
  }
  if(d.all["cswmIcoOn"+id])
  {
    d.all["cswmIco"+id].style.display="none";
    d.all["cswmIcoOn"+id].style.display="inline";
  }
  d.all["cswmItem"+id].style.color="#000000";
  d.all["cswmExpand"+id].style.color="#000000";
  if(bgco)
  {
    d.all["cswmItem"+id].style.backgroundColor=bgco;
    d.all["cswmExpand"+id].style.backgroundColor=bgco;
  }
  else
  {
    d.all["cswmItem"+id].style.backgroundColor="#B6BDD2";
    d.all["cswmExpand"+id].style.backgroundColor="#B6BDD2";
  }
  if(d.all["cswmExpandIc"+id])
  {
    d.all["cswmExpandIc"+id].src=cswmExdIS.src;
  }
  cswmNHM(id,l);
  cswmNH[l-1]=id;
}

function cswmNHM(id,l)
{
  if(cswmNH[l-1]!=id)
  {
    var count=l-1;
    for(count=l-1;count<cswmNH.length;count++)
    {
      cswmDiI(cswmNH[count]);
    }
    cswmNH.length=l;
  }
}

function cswmDiI(id)
{
  var d;
  d=document;
  if(!d.all["cswmItem"+id])
  {
    return;
  }
  var bgco;
  try
  {
    bgco = d.all["cswmItem"+id].getAttribute('cswmUnSelColor');
  }
  catch(e)
  {
    bgco = false;
  }
  if(d.all["cswmIcoOn"+id])
  {
    d.all["cswmIco"+id].style.display="inline";
    d.all["cswmIcoOn"+id].style.display="none";
  }
  d.all["cswmItem"+id].style.color="#000000";
  d.all["cswmExpand"+id].style.color="#000000";
  if(bgco)
  {
    d.all["cswmItem"+id].style.backgroundColor=bgco;
    d.all["cswmExpand"+id].style.backgroundColor=bgco;
  }
  else
  {
    d.all["cswmItem"+id].style.backgroundColor="";
    d.all["cswmExpand"+id].style.backgroundColor="";
  }
  if(d.all["cswmExpandIc"+id])
  {
    d.all["cswmExpandIc"+id].src=cswmExIS.src;
  }
}

function cswmHideSelectBox(boolHide,arrSelectList)
{
}

function cswmSP(id,itemid)
{
  if(!itemid)
  {
    if(cswmFP)
    {
      cswmSEL=cswmXOff;
      cswmSET=cswmYOff;
      cswmSEH=1;
      cswmSEW=1;
      cswmFP=0;
    }
    else
    {
      if(!document.all[cswmSI])
      {
        return;
      }
      cswmSE=new Object(document.all[cswmSI]);
      var cswmPrO=cswmSE;
      var cswmPrT="";
      cswmSEL=cswmSE.offsetLeft+cswmXOff;
      cswmSET=cswmSE.offsetTop+cswmYOff;
      cswmSEH=cswmSE.offsetHeight;
      cswmSEW=cswmSE.offsetWidth;
      while(cswmPrT!="BODY")
      {
        cswmPrO=cswmPrO.offsetParent;
        cswmSEL+=cswmPrO.offsetLeft;
        cswmSET+=cswmPrO.offsetTop;
        cswmPrT=cswmPrO.tagName;
      }
    }
    document.all["cswmPopup"+id].style.display="block";
    cswmPW=document.all["cswmPopup"+id].clientWidth;
    cswmPH=document.all["cswmPopup"+id].clientHeight;
    cswmBW=document.body.clientWidth;
    cswmBH=document.body.clientHeight;
    cswmSLA=document.body.scrollLeft;
    cswmSTA=document.body.scrollTop;
    switch(cswmSPnt)
    {
      case "above":
  	    cswmPx[cswmPx.length]=cswmSEL;
        cswmPy[cswmPy.length]=cswmSET-cswmPH;
        cswmCA();
        cswmCR();
      break;
      case "below":
        cswmPx[cswmPx.length]=cswmSEL;
        cswmPy[cswmPy.length]=cswmSET+cswmSEH;
        cswmCB();
        cswmCR();
      break;
      case "right":
	    cswmPx[cswmPx.length]=cswmSEL+cswmSEW;
        cswmPy[cswmPy.length]=cswmSET;
        cswmCR();
        cswmCB();
      break;
      case "left":
	    cswmPx[cswmPx.length]=cswmSEL-cswmPW;
        cswmPy[cswmPy.length]=cswmSET;
        cswmCL();
        cswmCB();
        cswmDir="left";
      break;
    }
    cswmXOff=0;
    cswmYOff=0;
    document.all["cswmPopup"+id].style.left=cswmPx[cswmPx.length-1];
    document.all["cswmPopup"+id].style.top=cswmPy[cswmPy.length-1];
    cswmPI[cswmPI.length]=id;
  }
  else
  {
    var d;
    d=document;
    cswmPx[cswmPx.length]=document.all["cswmPopup"+cswmPI[cswmPI.length-1]].clientWidth+cswmPx[cswmPx.length-1]-4;
    var szPrE="";
    if(d.all["cswmItem"+itemid].parentElement.offsetTop==0)
    {
      if(navigator.platform=="MacPPC")
      {
        var szPrE="parentElement.parentElement.";
      }
      else
	    if(d.all["cswmItem"+itemid].parentElement.parentElement.parentElement.parentElement.className!="cswmPopupBox")
        {
          var szPrE="parentElement.parentElement.parentElement.";
        }
    }
    cswmPy[cswmPy.length]=eval("d.all[\"cswmItem"+itemid+"\"].parentElement."+szPrE+"offsetTop")+cswmPy[cswmPy.length-1];
    document.all["cswmPopup"+id].style.display="block";
    cswmPW=document.all["cswmPopup"+id].clientWidth;
    cswmPH=document.all["cswmPopup"+id].clientHeight;
    var cswmPrW=document.all["cswmPopup"+cswmPI[cswmPI.length-1]].clientWidth;
    cswmAR=cswmBW-cswmPx[cswmPx.length-1]+cswmSLA;
    cswmAB=cswmBH-cswmPy[cswmPy.length-1]+cswmSTA;
    if(cswmPx[cswmPx.length-2]==cswmSLA)
    {
      cswmDir="right";
    }
    if((cswmAR<cswmPW)||(cswmDir=="left"))
    {
      cswmMB=(cswmPx[cswmPx.length-1]-cswmPW-cswmPrW)+8;
      if((cswmMB>=0)&&(cswmMB>cswmSLA))
      {
        cswmDir="left";
      }
      else
      {
        cswmMB=cswmSLA;
      }
      cswmPx[cswmPx.length-1]=cswmMB;
    }
    if(cswmAB<cswmPH)
    {
      cswmMB=cswmPy[cswmPy.length-1]-(cswmPH-cswmAB);
      if(cswmMB<0)
      {
        cswmMB=cswmSTA;
      }
      cswmPy[cswmPy.length-1]=cswmMB;
    }
    document.all["cswmPopup"+id].style.left=cswmPx[cswmPx.length-1];
    document.all["cswmPopup"+id].style.top=cswmPy[cswmPy.length-1];
    cswmPI[cswmPI.length]=id;
  }
  if(navigator.platform!="MacPPC")
  {
    cswmMS(id,cswmPx[cswmPx.length-1],cswmPy[cswmPy.length-1],cswmPW,cswmPH);
  }
  if(navigator.platform!='MacPPC')
  {
    if(navigator.userAgent.indexOf('MSIE 5.0')<=0)
    {
      cswmIFSH(id);
    }
  }
}

function cswmHP(level)
{
  if(cswmClkd==-1)
  {
    return false;
  }
  else 
    if(level==0)
    {
      cswmClkd=-1;
      var id = cswmPI[0];
      var count=0;
      for(count=0;count<cswmNH.length;count++)
      {
        cswmDiI(cswmNH[count]);
      }
      cswmNH.length=0;
      cswmButtonNormal("cswmMenuButton"+id);
      cswmButtonClickState=false;
    }
    var count=level;
    for(count=level;count<cswmPI.length;count++)
    {
      document.all["cswmPopup"+cswmPI[count]].style.display="none";
      if(document.all['cswmIFrame'+cswmPI[count]])
      {
        document.all['cswmIFrame'+cswmPI[count]].style.display='none';
      }
    }
    cswmPI.length=level;
    cswmPx.length=level;
    cswmPy.length=level;
    if(navigator.platform!="MacPPC")
    {
      cswmDS(level);
    }
}

function cswmIFSH(id)
{
  if(document.readyState!='complete')
  {
    return false;
  }
  var ifr;
  if(!document.all['cswmIFrame'+id])
  {
    ifr="<iframe src=\"javascript:void 0;\" id=\"cswmIFrame" + id + "\" scrolling=\"no\" frameborder=\"0\" style=\"position:absolute;top:0x;left:0px;z-index:998;display:none\"></iframe>";
    document.body.insertAdjacentHTML('beforeEnd',ifr);
  }
  if(document.all['cswmIFrame'+id])
  {
    ifr=document.all['cswmIFrame'+id].style;
    ifr.top=cswmPy[cswmPy.length-1];
    ifr.left=cswmPx[cswmPx.length-1];
    ifr.width=cswmPW;
    ifr.height=cswmPH;
    ifr.filter='progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)';
    ifr.display='block';
  }
}

function cswmCR()
{
  cswmAR=(cswmBW+cswmSLA)-cswmPx[cswmPx.length-1];
  if(cswmAR<cswmPW+4)
  {
    if(cswmSPnt=="below"||cswmSPnt=="above")
    {
      cswmMB=cswmPx[cswmPx.length-1]-(cswmPW-cswmAR)-4;
      if(cswmMB<0||cswmMB<cswmSLA)
      {
        cswmMB=cswmSLA;
      }
      cswmPx[cswmPx.length-1]=cswmMB;
    }
    else
    {
      cswmMB=cswmSEL-cswmPW;
      if(cswmMB>=0)
      {
        cswmPx[cswmPx.length-1]=cswmMB;
      }
    }
  }
}

function cswmCL()
{
  if(cswmPx[cswmPx.length-1]<(cswmSLA))
  {
    cswmPx[cswmPx.length-1]=cswmSEL+cswmSEW;
    cswmCR();
  }
}

function cswmCB()
{
  cswmAB=(cswmBH+cswmSTA)-cswmPy[cswmPy.length-1];
  if(cswmAB<cswmPH)
  {
    if(cswmSPnt=="below")
    {
      cswmMB=cswmPy[cswmPy.length-1]-cswmPH-cswmSEH;
      if(cswmMB>=0)
      {
        cswmPy[cswmPy.length-1]=cswmMB;
      }
    }
    else
    {
      cswmMB=cswmPy[cswmPy.length-1]-(cswmPH-cswmAB);
      if(cswmMB<0||cswmMB<cswmSTA)
      {
        cswmMB=cswmSTA;
      }
      cswmPy[cswmPy.length-1]=cswmMB;
    }
  }
}

function cswmCA()
{
  if(cswmPy[cswmPy.length-1]<(cswmSTA))
  {
    cswmPy[cswmPy.length-1]=cswmSET+cswmSEH;
    cswmCB();
  }
}

function cswmMS(id,x,y,w,h)
{
  var rt;
  var rs;
  var i;
  for (i=0;i<4;i++)
  {
    rt=document.createElement("div");
    rs=rt.style;
    rs.position="absolute";
    rs.zIndex=1999;
    rs.left=(x+i)+(4);
    rs.top=((y+3)-i)+(4);
    if(cswmPW>8)
    {
      rs.width=w-(i*2);
    }
    if(cswmPH>8)
    {
      rs.height=(h-6)+(i*2);
    }
    rs.backgroundColor="#000000";
    rs.filter="alpha(opacity=6)";
    document.all["cswmPopup"+cswmPI[cswmPI.length-1]].insertAdjacentElement("beforeBegin",rt);
    cswmSdw[cswmSdw.length]=rt;
  }
}

function cswmDS(level)
{
  var i;
  var Keep=level*4;
  for(i=Keep;i<cswmSdw.length;i++)
  {
    cswmSdw[i].removeNode(true);
  }
  cswmSdw.length=Keep;
}

function cswmShowInFrame(MenuID,x,y)
{
  x+=document.body.scrollLeft;
  y+=document.body.scrollTop;
  cswmShow(MenuID,'','below',x,y,1);
}

function cswmRefresh()
{
  if(navigator.platform=="MacPPC")
  {
    cswmResize();
  }
}

var cswmButtonClickState=false;
var cswmCurrentButtonId;
var cswmButtonsObj;
var cswmNeedPosInit=true;
var cswmDrag=false;
var cswmDragX;
var cswmDragY;
var cswmOnMouseMove="";
var cswmOnMouseUp="";
var cswmBodyCursor="";
var cswmDockObj;
var cswmIsDock=true;  
var cswmDockSpace="";
var cswmTop=0;
var cswmLeft=0;
var cswmMBIF;
var cswmMBIFT;

function cswmButtonDown(id,gid)
{
  cswmCurrentButtonId=id;
  cswmButtonSunken(id);
  if(cswmIsDock)
  {
    cswmShow(gid, id, 'below', 1, 1);
  }
  else
  {
    cswmShow(gid, id, 'below', 2, 2);
  }
}

function cswmButtonSelect(id,gid)
{
  if(!cswmButtonClickState)
  {
    cswmButtonRaised(id);
  }
  else
  {
    cswmButtonNormal(cswmCurrentButtonId);
    clearTimeout(cswmTI);
    cswmButtonDown(id,gid);
  }
}

function cswmButtonUnSelect(id)
{
  if(!cswmButtonClickState)
  {
    cswmButtonNormal(id);
  }
  else
  {
    cswmHide();
  }
}

function cswmButtonRaised(id)
{
  var obj = document.all(id).style;
  obj.borderTopColor = "#0A246A";
  obj.borderLeftColor = "#0A246A";
  obj.borderBottomColor = "#0A246A";
  obj.borderRightColor = "#0A246A";
  obj.backgroundColor = "#B6BDD2";
  obj.paddingBottom = "4px";
  obj.paddingTop = "4px";
  obj.paddingLeft = "6px";
  obj.paddingRight = "6px";
  obj.color = "#000000";
}

function cswmButtonSunken(id)
{
  var obj = document.all(id).style;
  obj.borderTopColor = "#808080";
  obj.borderLeftColor = "#808080";
  obj.borderBottomColor = "#ffffff";
  obj.borderRightColor = "#ffffff";
  obj.backgroundColor = "#d4d0c8";
  obj.paddingBottom = "3px";
  obj.paddingTop = "5px";
  obj.paddingLeft = "7px";
  obj.paddingRight = "5px";
  obj.color = "#000000";
}

function cswmButtonNormal(id)
{
  var obj = document.all(id).style;
  obj.borderTopColor = "#d4d0c8";
  obj.borderLeftColor = "#d4d0c8";
  obj.borderBottomColor = "#d4d0c8";
  obj.borderRightColor = "#d4d0c8";
  obj.backgroundColor = "#d4d0c8";
  obj.paddingBottom = "4px";
  obj.paddingTop = "4px";
  obj.paddingLeft = "6px";
  obj.paddingRight = "6px";
  obj.color = "#000000";
}

function cswmMenuBarPos()
{
  cswmButtonsObj = document.all.cswmButtons;
  cswmDockObj = document.all.cswmDock;
  cswmButtonsObj.style.left = cswmLeft;
  cswmButtonsObj.style.top = cswmTop;
  cswmNeedPosInit = false;
  cswmDockObj.style.height = cswmButtonsObj.offsetHeight;
  document.all.cswmInrDck.style.height = cswmButtonsObj.offsetHeight-2;
  if(navigator.platform!="MacPPC")
  {
    cswmMBIF=document.all.cswmMenuBarIFrame.style;
    cswmMBIF.left=cswmLeft;
    cswmMBIF.top=cswmTop;
    cswmMBIF.width=cswmButtonsObj.offsetWidth;
    cswmMBIF.height=cswmButtonsObj.offsetHeight;
    if(navigator.userAgent.indexOf("MSIE 5.0")<=0)
    {
      cswmMBIF.display='block';
    }
  }
}

function cswmBarDragStart()
{
  if(cswmNeedPosInit)
  {
    cswmMenuBarPos();
  }
  if(!cswmDrag)
  {
    cswmDrag = true;
    cswmDragX = window.event.x;
    cswmDragY = window.event.y;
    cswmBodyCursor = document.body.style.cursor;
    document.body.style.cursor = "move";
    cswmButtonsObj.style.cursor = "move";
    cswmOnMouseMove = document.onmousemove;
    document.onmousemove = cswmBarDrag;
    cswmOnMouseUp = document.body.onmouseup;
    document.onmouseup=cswmBarDragEnd;
    if(cswmCSDS)
    {
      if(!cswmMBZ)
      {
        _csdsZIndexArray[_csdsZIndexArray.length]=cswmButtonsObj.style;
        document.all.cswmMenuBarIFrame.style.zIndex = 50;
        cswmMBZ=true;
      }
      var count=0;
      for(count=0;count<_csdsZIndexArray.length;count++)
      {
        _csdsZIndexArray[count].zIndex-=1;
      }
      cswmButtonsObj.style.zIndex=_csdsZIndex + 1;
    }
  }
}

function cswmBarDragEnd()
{
  if (cswmDrag)
  {
    cswmDrag = false;
    document.body.style.cursor = cswmBodyCursor;
    cswmButtonsObj.style.cursor = "default";
    document.onmousemove = cswmOnMouseMove;
    document.onmouseup = cswmOnMouseUp;
    if(cswmDockSpace != "")
    {
      cswmDock();
    }
  }
}

function cswmBarDrag()
{
  if(cswmIsDock)
  {
    cswmUnDock();
  }
  if(cswmDrag)
  {
    var csdsAdjustTop=0;
    var csdsAdjustBottom=0;
    if(cswmCSDS)
    {
      csdsAdjustTop=_csdsTop;
      csdsAdjustBottom=_csdsBottom;
    }
    if(window.event.x < cswmDragX)
    {
      cswmLeft -= cswmDragX - window.event.x;
      cswmDragX -= cswmDragX - window.event.x;
      cswmButtonsObj.style.left = cswmLeft + document.body.scrollLeft;
      if(navigator.platform!="MacPPC")
      {
        cswmMBIF.left=cswmButtonsObj.style.left;
      }
    }
    else
	  if(window.event.x > cswmDragX)
      {
        cswmLeft += window.event.x - cswmDragX;
        cswmDragX += window.event.x - cswmDragX;
        cswmButtonsObj.style.left = cswmLeft + document.body.scrollLeft;
        if(navigator.platform!="MacPPC")
        {
          cswmMBIF.left=cswmButtonsObj.style.left;
        }
      }
    if(window.event.y < cswmDragY)
    {
      cswmTop -= cswmDragY - window.event.y;
      cswmDragY -= cswmDragY - window.event.y;
      cswmButtonsObj.style.top = cswmTop + document.body.scrollTop;
      if(navigator.platform!="MacPPC")
      {
        cswmMBIF.top=cswmButtonsObj.style.top;
      }
    }
    else
	  if(window.event.y > cswmDragY)
      {
        cswmTop += window.event.y - cswmDragY;
        cswmDragY += window.event.y - cswmDragY;
        cswmButtonsObj.style.top = cswmTop + document.body.scrollTop;
        if(navigator.platform!="MacPPC")
        {
          cswmMBIF.top=cswmButtonsObj.style.top;
        }
      }
    if(parseInt(cswmButtonsObj.style.top) < (5 + document.body.scrollTop + csdsAdjustTop))
    {
      cswmDockSpace = "top";
      cswmDisplayDock(true);
    }
    else
	  if((parseInt(cswmButtonsObj.style.top) + cswmButtonsObj.offsetHeight) > (document.body.clientHeight - (5-document.body.scrollTop) - csdsAdjustBottom))
      {
        cswmDockSpace = "bottom";
        cswmDisplayDock(true);
      }
      else
      {
        cswmDockSpace = "";
        cswmDisplayDock(false);
      }
  }
}

function cswmDock(location)
{
  cswmButtonsObj.style.borderWidth = "0px";
  cswmLeft = 0;
  cswmButtonsObj.style.left = cswmLeft + document.body.scrollLeft;
  cswmButtonsObj.style.width = "100%";
  cswmDisplayDock(false);
  if(String(location) != "undefined")
  {
    if(cswmNeedPosInit)
    {
      cswmMenuBarPos();
    }
    cswmDockSpace = location;
  }
  if(cswmDockSpace == "top")
  {
    var csdsAdjust=0;
    if(cswmCSDS)
    {
      csdsAdjust=_csdsTop;
    }
    cswmTop=0+csdsAdjust;
    cswmButtonsObj.style.top = cswmTop + document.body.scrollTop;
    document.body.style.paddingTop = parseInt(cswmDockObj.style.height)-2+csdsAdjust;
    if(cswmCSDS)
    {
      _csdsCoopDock('cswm','top',parseInt(cswmDockObj.style.height)-2);
    }
  }
  else
    if(cswmDockSpace == "bottom")
    {
      var csdsAdjust=0;
      if(cswmCSDS)
      {
        csdsAdjust=_csdsBottom;
      }
      cswmTop = (document.body.clientHeight - cswmButtonsObj.offsetHeight)-csdsAdjust;
      cswmButtonsObj.style.top = cswmTop + document.body.scrollTop;
      document.body.style.paddingBottom = parseInt(cswmDockObj.style.height)-2+csdsAdjust;
      if(cswmCSDS)
      {
        _csdsCoopDock('cswm','bottom',parseInt(cswmDockObj.style.height)-2);
      }
    }
  cswmIsDock = true;
  if(navigator.platform!="MacPPC")
  {
    cswmMBIF.left=cswmButtonsObj.style.left;
    cswmMBIF.top=cswmButtonsObj.style.top;
    cswmMBIF.width=cswmButtonsObj.offsetWidth;
    cswmMBIF.height=cswmButtonsObj.offsetHeight;
  }
}

function cswmUnDock()
{
  cswmDisplayDock(true);
  cswmButtonsObj.style.borderWidth = "1px";
  cswmButtonsObj.style.width = 1;
  cswmIsDock = false;
  cswmDockSpace = "";
  if(cswmCSDS)
  {
    _csdsCoopUnDock('cswm');
    document.body.style.paddingTop=_csdsTop;
    document.body.style.paddingBottom=_csdsBottom;
  }
  else
  {
    document.body.style.paddingTop=0;
    document.body.style.paddingBottom=0;
  }
  if(navigator.platform!="MacPPC")
  {
    cswmMBIF.left=cswmButtonsObj.style.left;
    cswmMBIF.top=cswmButtonsObj.style.top;
    cswmMBIF.width=cswmButtonsObj.offsetWidth;
    cswmMBIF.height=cswmButtonsObj.offsetHeight;
  }
}

function cswmDisplayDock(state)
{
  if(state)
  {
    if(cswmDockSpace == "top")
    {
      var csdsAdjust=0;
      if(cswmCSDS)
      {
        csdsAdjust=_csdsTop;
      }
      cswmDockObj.style.top = 0 + document.body.scrollTop + csdsAdjust;
    }
    else
    {
      var csdsAdjust=0;
      if(cswmCSDS)
      {
        csdsAdjust=_csdsBottom;
      }
      cswmDockObj.style.top = ((document.body.clientHeight - cswmDockObj.offsetHeight) + document.body.scrollTop) - csdsAdjust;
    }
    cswmDockObj.style.left = document.body.scrollLeft;
    cswmDockObj.style.display = "block";
  }
  else
  {
    cswmDockObj.style.display = "none";
  }
}

function cswmFloat()
{
  if(cswmNeedPosInit)
  {
    cswmMenuBarPos();
  }
  cswmButtonsObj.style.top = cswmTop + document.body.scrollTop;
  cswmButtonsObj.style.left = cswmLeft + document.body.scrollLeft;
  if(navigator.platform!="MacPPC")
  {
    if(navigator.userAgent.indexOf("MSIE 5.0")<=0)
    {
      clearTimeout(cswmMBIFT);
      cswmMBIF.display='none';
      cswmMBIF.left=cswmButtonsObj.style.left;
      cswmMBIF.top=cswmButtonsObj.style.top;
      cswmMBIFT=setTimeout("cswmMBIF.display='block';",100);
    }
  }
}

function cswmResize()
{
  if(cswmCSDS)
  {
    _csdsCoopResize();
    return;
  }
  if(cswmDockSpace == 'bottom')
  {
    cswmDock('bottom');
  }
}

function cswmMenuBarInit()
{
  if(typeof(_csds)!='undefined')
  {
    cswmCSDS=true;
    if(_csdsZIndex<999)
    {
      _csdsZIndex=999;
    }
  }
//  cswmMenuBarPos();
  this.attachEvent('onscroll', cswmFloat);
  this.attachEvent('onresize', cswmResize);
}

//-->

function DisableContextMenu(e){return false;}
document.oncontextmenu = DisableContextMenu;
</script>
<title><%= Session("Titulo")%>Supervisor - Ticket - Usuario: <%= l_username %></title>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" onLoad="cswmMenuBarInit();" onResize="cswmRefresh()">

<% 

dim linea, sql, rs, rs2, rs3, rs4, submenu, orden, orden2, padre, destino, perfil, username, menuraiz, f, id

function Habilitado(perfil,perfilmenu)
   Dim arreglo
   Dim i
   Dim autorizado
   
   if (perfilmenu = "*") then
       autorizado = true   
   else
	   if (UCase(perfilmenu) = UCase(perfil)) then
		    autorizado = true
	   else
		    arreglo = split(perfilmenu,";")
			autorizado = false
			for i=0 to (UBound(arreglo))
			   if (UCase(arreglo(i)) = UCase(perfil)) then
			      autorizado = true
			   end if
			next
	   end if
	end if
   Habilitado = autorizado
end function

function habilitadoPermiso(menu, usuario)
	l_planta = 1
	habilitadoPermiso = false
	'Busco el menú, usuario y planta
	sql = "SELECT menunro FROM tkt_usu_men_pla "
	sql = sql & " WHERE iduser = '" & usuario & "'"
	sql = sql & " AND menunro = " & menu
	sql = sql & " AND planro = " & l_planta
	rsOpen mirs, cn, sql, 0
	If Not mirs.EOF Then
	    habilitadoPermiso = True
	End If
	mirs.close
end function


Set menuraiz = CreateObject("Scripting.Dictionary")

Menu = ""
username = UCase(Session("Username"))
if not username = "SUPER" then
	sql = "SELECT perfnom FROM user_per"
	sql = sql & " inner join perf_usr on perf_usr.perfnro = user_per.perfnro"
	sql = sql & " where upper(iduser) = '" & username & "'"
	rs3.Maxrecords = 1
	rsOpen rs3, cn, sql, 0
	perfil = rs3("perfnom")
	rs3.Close
end if

sql = "SELECT * FROM menuraiz where upper(menuraiz.menudesc) = 'SUP'"
rs4.Maxrecords = 1
rsOpen rs4, cn, sql, 0
raiz = rs4("menunro")
rs4.Close

sql = "SELECT MenuName, MenuOrder, MenuRaiz, Parent, tipo, action, menuaccess, menuimg, '0',menunro FROM menumstr where menuraiz = " & raiz
'if not username = "SUPER" then
'	sql = sql & " AND (menuaccess LIKE '%" & perfil & "%' OR "
'	sql = sql & " menuaccess = '*') "
'end if
sql = sql & " ORDER BY parent desc, menuorder"
rsOpen rs, cn, sql, 0
menumstr = rs.GetRows
rs.Close

for i = 0 to Ubound(menumstr,2)

	menumstr(8,i) = menumstr(1,i) 'Lo copio para mantener el valor del orden original

	menumstr(3,i) = Left(menumstr(3,i),len(menumstr(3,i))-3)

	if isnull(menumstr(7,i)) or trim(menumstr(7,i)) = "" then
	  menumstr(7,i) = "/trivoliSwimming/shared/images/blank.gif"
	else
		menumstr(7,i) = "/trivoliSwimming/shared/images/" & menumstr(7,i)
	end if
	
	if menumstr(5,i) = "" then
	  menumstr(5,i) = "#"
	end if
	if menumstr(3,i) = "" then
	  menuraiz.add TRIM(menumstr(1,i)), left(menumstr(0,i),5)
	  menumstr(1,i) = Left(menumstr(0,i),5)
	  menumstr(2,i) = "0"
	else
	  if menuraiz.Exists(TRIM(menumstr(3,i))) then
	    menumstr(3,i) = menuraiz.item(TRIM(menumstr(3,i)))
  	    menumstr(2,i) = "1"
	  else
	    if len(menumstr(3,i)) = 1 then
 	      menumstr(3,i) = "  " & (menumstr(3,i))
		else  
 	      if len(menumstr(3,i)) = 2 then
 	        menumstr(3,i) = " " & (menumstr(3,i))
		  end if
		end if
	    menumstr(2,i) = "2"
	  end if
	end if
next

for i = 0 to Ubound(menumstr,2)
  if menumstr(2,i) = "2" then
    for j = 0 to Ubound(menumstr,2)
	  if trim(menumstr(3,i)) = trim(menumstr(1,j)) then
	    menumstr(2,i) = cstr(cint(menumstr(2,j)) + 1)
	  end if
	next
  end if
next


for i = 0 to Ubound(menumstr,2) - 1
  for j = i + 1 to Ubound(menumstr,2)
	if ((menumstr(2,i)) > (menumstr(2,j))) or ((menumstr(2,i) = menumstr(2,j)) and (menumstr(3,i) > menumstr(3,j)) or ((menumstr(2,i) = menumstr(2,j)) and (menumstr(3,i) = menumstr(3,j)) and (menumstr(8,i) > menumstr(8,j)))) then
	  for k = 0 to Ubound(menumstr,1)
	    auxi = menumstr(k,i)
	    menumstr(k,i) = menumstr(k,j)
		menumstr(k,j) = auxi
	  next
	end if
  next
next

%>	
<% 
i = 0 
padre = ""
while i <= Ubound(menumstr,2)
  if not isNumeric(trim(menumstr(3,i))) and (menumstr(3,i) <> "") then
    padre = trim(menumstr(3,i))
%>					
	<div id="cswmPopup<%= padre %>" class="cswmPopupBox" onselectstart="return false;">
		<table background="/trivoliSwimming/shared/images/background.gif" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<div style="border-style:solid; border-width: 1px; border-color:#666666 #666666 #666666 #666666">
						<div style="border-style:solid; border-width: 1px; border-color:#F9F8F7 #F9F8F7 #F9F8F7 #F9F8F7">
							<table border="0" cellpadding="0" cellspacing="0">
<%
    Do while (i <= Ubound(menumstr,2))
	  if (trim(menumstr(3,i)) <> padre) then
	    exit do
	  end if
  	  id = padre & "_" & menumstr(1,i)
      if trim(menumstr(0,i)) = "rule" then
%>	  
								<tr onMouseOver="cswmT('off');cswmNHM('space',1);" onMouseOut="cswmT(350);"><td align="Left" bgcolor="#a5a6a6" colspan="2" height="1"><img border="0" alt="" src="/trivoliSwimming/shared/images/divider.gif"></td>
								</tr>
<%
	  else
'if trim(menumstr(0,i)) = "Periodos de GTI" then 

        if (not Habilitado(perfil,menumstr(6,i))) and (not habilitadoPermiso(menumstr(9,i),username)) then
'        if (not username = "SUPER") and (instr(menumstr(6,i), perfil) = 0) and (menumstr(6,i)<> "*") then
'response.write "<script>alert('"&menumstr(0,i)&"')</script>"
%>
								<tr onMouseOver="cswmT('off');cswmNHM('<%= id %>',1);cswmST(1);" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
									    <span class="cswmDisabled">
										<%= menumstr(0,i) %>
									    </span>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand">
<%
        else
        if trim(menumstr(4,i)) = "I" then
%>									
								<tr onMouseOver="cswmT('off');cswmHiI('<%= id %>',1);cswmST(1);" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);location.href='<%= Replace(menumstr(5,i), "'", "\'") %>';">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
										<%= menumstr(0,i) %>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand">
<%
        else
%>									
								<tr onMouseOver="cswmT('off');cswmHiI('<%= id %>',1);cswmST(1,<%= menumstr(1,i) %>,'<%= id %>');" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);location.href='<%= Replace(menumstr(5,i), "'", "\'") %>';">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
										<%= menumstr(0,i) %>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand" style="padding-right:5"><img id="cswmExpandIcFile_1" src="/trivoliSwimming/shared/images/Popup.gif" width="10" height="10" alt="" border="0">
<%
        end if
        end if
      end if
%>																		
									</td>
								</tr>
<%
      i = i + 1
    Loop
%>
							</table>
						</div>
					</div>
				</td>
			</tr>
		</table>
	</div>
<%	
  else
    i = i + 1  
  end if
Wend
%>

<% 
i = 0 
padre = ""
while i <= Ubound(menumstr,2)
  if not isNumeric(trim(menumstr(3,i))) or (menumstr(3,i) = "") then
    i = i + 1
  else
    padre = trim(menumstr(3,i))
%>					
	<div id="cswmPopup<%= padre %>" class="cswmPopupBox" onselectstart="return false;" onMouseOver="cswmHiI('<%= padre %>',<%= menumstr(2,i) %>);">
		<table background="/trivoliSwimming/shared/images/background.gif" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<div style="border-style:solid; border-width: 1px; border-color:#666666 #666666 #666666 #666666">
						<div style="border-style:solid; border-width: 1px; border-color:#F9F8F7 #F9F8F7 #F9F8F7 #F9F8F7">
							<table border="0" cellpadding="0" cellspacing="0">
<%
    do while (i <= Ubound(menumstr,2))
	  if (trim(menumstr(3,i)) <> padre) then
	    exit do
	  end if 
 	  id = padre & "_" & menumstr(1,i)
      if trim(menumstr(0,i)) = "rule" then
%>	  
								<tr onMouseOver="cswmT('off');cswmNHM('space',1);" onMouseOut="cswmT(350);"><td align="Left" bgcolor="#a5a6a6" colspan="2" height="1"><img border="0" alt="" src="/trivoliSwimming/shared/images/divider.gif"></td>
								</tr>
<%
	  else
        if (not username = "SUPER") and (instr(menumstr(6,i), perfil) = 0) and (not habilitadoPermiso(menumstr(9,i),username) ) and (menumstr(6,i)<> "*") then
%>
								<tr onMouseOver="cswmT('off');cswmNHM('<%= id %>',<%= menumstr(2,i) %>);cswmST(<%= menumstr(2,i) %>);" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
									    <span class="cswmDisabled">
										<%= menumstr(0,i) %>
									    </span>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand">	  
<%
	    else
          if trim(menumstr(4,i)) = "I" then
%>									
								<tr onMouseOver="cswmT('off');cswmHiI('<%= id %>',<%= menumstr(2,i) %>);cswmST(<%= menumstr(2,i) %>);" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);location.href='<%= Replace(menumstr(5,i), "'", "\'") %>';">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
										<%= menumstr(0,i) %>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand">
									</td>
								</tr>
<%
          else
%>									
								<tr onMouseOver="cswmT('off');cswmHiI('<%= id %>',<%= menumstr(2,i) %>);cswmST(<%= menumstr(2,i) %>,<%= menumstr(1,i) %>,'<%= id %>');" onMouseOut="cswmT(350);cswmST();" onClick="cswmHP(0);location.href='<%= Replace(menumstr(5,i), "'", "\'") %>';">
									<td nowrap bgcolor="" id="cswmItem<%= id %>" class="cswmItem">
									    <span><img align="absmiddle" id="cswmIco<%= id %>" src="<%= menumstr(7,i) %>" alt="" border="0" height="16" width="16"><img align="absmiddle" alt="" border="0" height="1" width="12" src="/trivoliSwimming/shared/images/ClearPixel.gif"></span>
										<%= menumstr(0,i) %>
									</td>
									<td bgcolor="" id="cswmExpand<%= id %>" class="cswmExpand" style="padding-right:5"><img id="cswmExpandIc<%= id %>" src="/trivoliSwimming/shared/images/Popup.gif" width="10" height="10" alt="" border="0">
									</td>
								</tr>
<%
          end if
        end if
      end if
      i = i + 1
    Loop
%>
							</table>
						</div>
					</div>
				</td>
			</tr>
		</table>
	</div>
<%	
  end if
Wend
%>
	<div class="cswmDock" id="cswmDock">
	    <span id="cswmInrDck" style="border:dotted #ffffff 1px; width:100%;"><img src="/trivoliSwimming/shared/images/ClearPixel.gif" width="1" height="1"></span>
	</div>
	<div class="cswmButtons" style="width:1px" id="cswmButtons" onselectstart="return false;">
		<div class="cswmInnerBorder">
			<table width="1" cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td style="padding:2px; height:100%" nowrap>
<!--  Quitado para que no muestre la opcion de mover el menu

					    <span class="cswmHandle" id="cswmHandle" onMouseDown="cswmBarDragStart(document.all.cswmButtons);"></span>
					    <span style="width:2px"></span>
 -->
 					</td>
<% 
i = 0 
while i <= Ubound(menumstr,2)
	if menumstr(3,i) = "" then
       if (not username = "SUPER") and (instr(menumstr(6,i), perfil) = 0) and (menumstr(6,i)<> "*" ) and (not habilitadoPermiso(menumstr(9,i),username)) then
%>
					<td class="cswmButton" id="cswmMenuButton<%= menumstr(1,i) %>" onMouseOut="cswmButtonUnSelect(this.id);"  nowrap>	
				    <span class="cswmDisabled">
					<%= menumstr(0,i) %>
				    </span>
<% Else %>
					<td class="cswmButton" id="cswmMenuButton<%= menumstr(1,i) %>" onMouseOver="cswmButtonSelect(this.id, '<%= menumstr(1,i) %>');" onMouseOut="cswmButtonUnSelect(this.id);" onMouseDown="cswmButtonDown(this.id, '<%= menumstr(1,i) %>');" nowrap>	
					<%= menumstr(0,i) %>
<% End If %>					
					</td>
<%
	end if
  i = i + 1
Wend
%>
					<td width="100%">
					</td>
				</tr>
			</table>
		</div>
	</div>
	<iframe src="javascript:void 0;" id="cswmMenuBarIFrame" scrolling="no" frameborder="0" style="position:absolute;top:0px;left:0px;z-index:998;display:none"></iframe>
<div>
<iframe src="principal.asp" frameborder="0" width="100%" height="100%" scrolling="no" style="position:absolute;top:25"></iframe>
</div>
<script>cswmMenuBarPos();cswmDock('top');</script>
</body>
</html>
