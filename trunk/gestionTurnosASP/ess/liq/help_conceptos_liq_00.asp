<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
Archivo: help_conceptos_liq_00.asp
Descripción: Ayuda de conceptos
Autor : FFavre
Fecha: 10/2003
Modificado:
	25-11-03 FFavre Se agregaron periodos retroactivos.
	04-02-04 FFavre Actualiza la cantidad de decimales definidos para el concepto en la vent. llamadora.
    05-10-04 - Scarpa D. - Correccion de las novedades retroactivas		
	25-10-05 - Leticia A. - Adecuacion a Autogestion 
-->
<% 
' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma
 
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
 
' Filtro
  l_etiquetas = "C&oacute;digo:;Descripci&oacute;n:;F&oacute;rmula:;Par&aacute;metro:"
  l_Campos    = "concepto.conccod;concepto.concabr;formula.fordabr;tipopar.tpadabr"
  l_Tipos     = "T;T;T;T"
 
' Orden
  l_Orden     = "C&oacute;digo:;Descripci&oacute;n:;F&oacute;rmula:;Par&aacute;metro:"
  l_CamposOr  = "concepto.conccod;concepto.concabr;formula.fordabr;tipopar.tpadabr"
 
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Conceptos de carga individual - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
	abrirVentana('/ess/ess/shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('/ess/ess/shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function Pasar_Valor(){
	if (document.ifrm.datos.cabnro.value == "0")
		alert("No selecciono ningun concepto");
	else{
		opener.document.all.concnro.value = document.ifrm.datos.cabnro.value;
		opener.document.all.conccod.value = document.ifrm.datos.conccod.value;
		opener.document.all.concabr.value = document.ifrm.datos.concabr.value;
		opener.document.all.tpanro.value  = document.ifrm.datos.tpanro.value;
		opener.document.all.tpadabr.value = document.ifrm.datos.tpadabr.value;
		opener.document.all.unidesc.value = document.ifrm.datos.unisigla.value;
		opener.document.all.conccantdec.value = document.ifrm.datos.conccantdec.value;
		opener.actualizarConcRetro(document.ifrm.datos.concretro.value);
		opener.document.all.nevalor.focus();
		opener.document.all.nevalor.select() ;
		window.close();
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
	<table border="0" cellpadding="0" cellspacing="0" height="100%">
    	<tr style="border-color :CadetBlue;">
        	<th align="left" class="th2">Conceptos de carga individual</th>
          	<th nowrap style= " text-align : right;">
		  		<a class=sidebtnSHW href="Javascript:orden('/ess/ess/liq/help_conceptos_liq_01.asp');">Orden</a>
				<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/liq/help_conceptos_liq_01.asp')">Filtro</a>
		  		&nbsp;&nbsp;
			</th>
        </tr>
        <tr valign="top" height="100%">
        	<td colspan="2" style="">
      	  		<iframe name="ifrm" src="help_conceptos_liq_01.asp" width="100%" height="100%"></iframe> 
			</td>
        </tr>
		<tr>
		    <td colspan="2" align="right" class="th2">
				<a class=sidebtnABM href="Javascript:Pasar_Valor();">Aceptar</a>&nbsp;
				<a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>&nbsp;
				<br> &nbsp;
			</td>
		</tr>
      </table>
</body>
</html>
