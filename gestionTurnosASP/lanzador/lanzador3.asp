<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= Session("Titulo")%>Ticket</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/serviciolocal/shared/css/tables4.css" rel="StyleSheet" type="text/css">
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
function tip(texto){
	document.all.titulo.value = texto;
}
function help(){
	abrirVentana(document.FormVar.ayuda.value,'',750,550);
}
function web(){
	abrirVentana(document.FormVar.inicio.value,'',750,550);
}
function cancel(){
	window.history.back();
}
function vertip(texto,objeto){
	document.all.tooltiptext.value = texto;
	document.all.tooltiptext.size = document.all.tooltiptext.value.length + 1;
	document.all.tooltip.style.left = event.x - (parseInt(document.all.tooltiptext.size))*5;
	document.all.tooltip.style.left = parseInt(objeto.style.left) - (objeto.width/2) - document.all.tooltiptext.value.length *3;
	document.all.tooltip.style.top = parseInt(objeto.style.top) - 25
	avertip();
}
function movtip(objeto){
	document.all.tooltiptext.size = document.all.tooltiptext.value.length + 1;
	document.all.tooltip.style.left = event.x - (parseInt(document.all.tooltiptext.size))*5;
	document.all.tooltip.style.left = parseInt(objeto.style.left) - (objeto.width/2) - document.all.tooltiptext.value.length *3;
	document.all.tooltip.style.top = parseInt(objeto.style.top) - 25
//	avertip();
}
var valor = 0;
var valor2;
function avertip(){
	valor = valor + 50
	document.all.tooltip.style.filter = 'alpha(opacity='+ valor +')';
	if (valor < 100){
		setTimeout('avertip()', 10);
	}else{
		document.all.tooltip.style.filter = 'alpha(opacity=100)';
	}
}
function novertip(){
	valor = valor - 50
	document.all.tooltip.style.filter = 'alpha(opacity='+ valor +')';
	if (valor > 0){
		setTimeout('novertip()', 10);
	}else{
		document.all.tooltip.style.left = 5000;
		document.all.tooltip.style.top = 5000;
		document.all.tooltip.style.filter = 'alpha(opacity=0)';
	}
}
function inicio(){
	document.all.ifrm2.src = "lanzador.asp?menu=html&ventana=03"
	window.document.title += " - <%= SESSION("USERNAME") %>";
}
function ventanas(dir,name){
	abrirVentanaCent(dir,name,600,440);
}

function logout(arg){
	abrirVentanaH('logout.asp?arg=1','',10,10);
}
</script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.all.ifrm.src='../shared/db/modulos.asp?menu=html';inicio();" onunload="Javascript://logout(1);">
<table width="502" height="359" border="0" cellpadding="0" cellspacing="0" background="">
  <tr> 
    <td> 
		<form name="FormVar" method="post">
		<input type="hidden" name="inicio" value="">
		<input type="hidden" name="email2" value="">
		<input type="hidden" name="ayuda" value="">
		<input type="hidden" name="empresa" value="">
		<input type="hidden" name="conexion" value="">
		<input type="hidden" name="debug" value=0>
		<input type="hidden" name="seguridad" value=0>
      </form>
	</td>
  </tr>
</table>
<iframe name="ifrm" style="visibility=hidden;" width="0" height="0"></iframe> <!--  -->
<iframe name="ifrm2" style="visibility=hidden;" src="" width="0" height="0"></iframe> 		
<iframe name="ifrm3" style="visibility=hidden;" src="/serviciolocal/shared/asp/timer.asp?menu=no" width="0" height="0"></iframe>
<div style="position:absolute; top:21px; left:31px;"> 
	<input class="blanc" readonly type="Text" align="absmiddle" name="titulo" size="40" border="0" style="border:none;width:200px;background-color :Transparent;color:#3377FF;font:18">
</div> 
<!--filter: alpha(opacity=50);-->
<table id="tooltip" cellpadding="0" cellspacing="0" align="center" style="position:absolute;filter: alpha(opacity=0); left: 163px; top: 351px; z-index:100; width: 176px; height: 27px;width:0;height:0;zoom:0.8">
  <tr>
		<td align="center"  valign="bottom" >
			<img src="../shared/images/tooltip_01.gif">
		</td>
		<td  align="center" valign="top" background="../shared/images/tooltip_02.gif" style="background-repeat : repeat-x">
			<input readonly border="0" id="tooltiptext" type="text" style="border:none;background-color:Transparent;color:#3377FF;font:18;heigth:2;text-align:center;" value="">
		</td>
		<td align="center" valign="bottom">
			<img src="../shared/images/tooltip_03.gif">
		</td>
	</tr>
</table>
<!--<a id="aadmper" href="#" style=""> <img id="admper" src="../shared/images/admper.gif" border="0" style="Filter: Gray();position:absolute; left: 271px; top: 40px" onMouseOver="vertip('Administración de Personal',this)" onMouseMove="movtip(this)" onMouseOut="novertip()" > 
</a><a id="aanalisis" href="#"> <img id="analisis" src="../shared/images/analisis.gif" border="0" style="position:absolute; left: 361px; top: 106px;Filter: Gray();" onMouseOver="vertip('Análisis de Remuneraciones',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="aautogestion" href="#"> <img id="autogestion" src="../shared/images/autogestion.gif" border="0" style="position:absolute; left: 373px; top: 155px;Filter: Gray();" onMouseOver="vertip('Autogestión',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="acapacitacion" href="#"> <img id="capacitacion" src="../shared/images/capacitacion.gif" border="0" style="position:absolute; left: 359px; top: 204px;Filter: Gray();" onMouseOver="vertip('Capacitación',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="acompetencias" href="#"> <img id="competencias" src="../shared/images/competencias.gif" border="0" style="position:absolute; left: 270px; top: 272px;Filter: Gray();" onMouseOver="vertip('Gestión de Competencias',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="aempleos" href="#"> <img id="empleos" src="../shared/images/empleos.gif" border="0" style="position:absolute; left: 322px; top: 250px;filter: Gray();" onMouseOver="vertip('Empleos y Postulantes',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="aevaluacion" href="#"> <img id="evaluacion" src="../shared/images/evaluacion.gif" border="0" style="position:absolute; left: 215px; top: 272px;filter: Gray();" onMouseOver="vertip('Gestión del Desempeño',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="agti" href="#"> <img id="gti" src="../shared/images/gti.gif" border="0" style="position:absolute; left: 164px; top: 252px;filter: Gray();" onMouseOver="vertip('Gestión de Tiempos',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="aliquidacion" href="#"> <img id="liquidacion" src="../shared/images/liquidacion.gif" border="0" style="position:absolute; left: 127px; top: 206px;filter: Gray();" onMouseOver="vertip('Liquidación',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="apoliticas" href="#"> <img id="politicas" src="../shared/images/politicas.gif" border="0" style="position:absolute; left: 112px; top: 155px;filter: Gray();" onMouseOver="vertip('Politicas y Procedimientos',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="asalud" href="#"> <img id="salud" src="../shared/images/salud.gif" border="0" style="position:absolute; left: 125px; top: 106px;filter: Gray();" onMouseOver="vertip('Salud Ocupacional',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> -->
<!--<a id="abuques" href="#">
	<img id="buques" src="../shared/images/buques.gif" border="0" style="position:absolute; left: 175px; top: 260px;filter: Gray();" onMouseOver="vertip('buques',this)" onMouseMove="movtip(this)" onMouseOut="novertip()">
</a>

<a id="aalertas" href="#"> <img id="alertas" src="../shared/images/alertas.gif" border="0" style="position:absolute; left: 299px; top: 209px;Filter: Gray();" onMouseOver="vertip('Alertas',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a>-->

<a id="aporteria" href="#"> <img id="porteria" src="../shared/images/capacitacion.gif" border="0" style="position:absolute; left: 299px; top: 209px;Filter: Gray();" onMouseOver="vertip('Porteria',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> <map name="Map">
</a> <a id="asupervisor" href="#"> <img id="supervisor" src="../shared/images/supervisor.gif" border="0" style="position:absolute; left: 136px; top: 114px; filter: Gray(); width: 40px; height: 40px;" onMouseOver="vertip('Seguridad y Auditoria',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="agerencial" href="#"> <img id="gerencial" src="../shared/images/infoger.gif" border="0" style="position:absolute; left: 230px; top: 60px;filter: Gray();" onMouseOver="vertip('Información Gerencial',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="aconfiguracion" href="#"> <img id="configuracion" src="../shared/images/config.gif" width="46" height="43" border="0" style="position:absolute; left: 315px; top: 115px;filter: Gray();" onMouseOver="vertip('Configuración',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="astock" href="#"> <img id="stock" src="../shared/images/stock.gif" width="45" height="45" border="0" style="position:absolute; left: 164px; top: 206px;filter: Gray();" onMouseOver="vertip('Stock',this)" onMouseMove="movtip(this)" onMouseOut="novertip()"> 
</a> <a id="" class="sidebtnABM" href="Javascript:logout(1);window.location= 'lanzador2.asp'//window.history.back();" style="position:absolute; left: 237px; top: 290px;">Salir</a> 


  <area shape="circle" coords="387,314,19" href="#" onClick="Javascript:logout(1);window.location= 'lanzador2.asp'//window.history.back();"  onMouseOver="tip('Volver');" onMouseOut="tip('');">
  <area id="email" shape="circle" coords="123,316,18" href="#" onMouseOver="tip('email');" onMouseOut="tip('');">
  <area id="home" shape="circle" coords="80,308,18" href="#" onClick="web();" onMouseOver="tip('Home');" onMouseOut="tip('');">
  <area id="ayuda" shape="circle" coords="63,265,19" href="javascript:help();"  onMouseOver="tip('Ayuda');" onMouseOut="tip('');">
</map>
<script>
document.all.tooltip.style.left = 5000;
document.all.tooltip.style.top = 5000;
document.all.tooltip.style.filter = 'alpha(opacity=0)';
</script>
</body>
</html>
