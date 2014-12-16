<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
Archivo: help_parametros_liq_00.asp
Descripción: Ayuda de parametros
Autor : FFavre
Fecha: 10/2003
Modificado:
	25-11-03 FFavre Falta ortográfica
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
  l_etiquetas = "C&oacute;digo:;Descripci&oacute;n:"
  l_Campos    = "tipopar.tpanro;tipopar.tpadabr"
  l_Tipos     = "N;T"
 
' Orden
  l_Orden     = "C&oacute;digo:;Descripci&oacute;n:"
  l_CamposOr  = "tipopar.tpanro;tipopar.tpadabr"
 
 Dim l_concnro
 
 l_concnro = request("concnro")
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Par&aacute;metros del conceptos - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
	abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function Pasar_Valor(){
	if (document.ifrm.datos.cabnro.value == "0")
		alert("No selecciono ningun parámetro");
	else{
		opener.document.all.tpanro.value  = document.ifrm.datos.cabnro.value;
		opener.document.all.tpadabr.value = document.ifrm.datos.tpadabr.value;
		opener.document.all.unidesc.value = document.ifrm.datos.unisigla.value;
		opener.document.all.nevalor.focus();
		opener.document.all.nevalor.select() ;
		window.close();
	}
}

function param(){
	return('concnro=<%= l_concnro%>')
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
	<table border="0" cellpadding="0" cellspacing="0" height="100%">
    	<tr style="border-color :CadetBlue;">
        	<th align="left" class="th2">Par&aacute;metros del conceptos</th>
          	<th nowrap style=" text-align:right;" class="th2">
		  		<a class=sidebtnSHW href="Javascript:orden('/ess/ess/liq/help_parametros_liq_01.asp');">Orden</a>
				<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/liq/help_parametros_liq_01.asp')">Filtro</a>
		  		&nbsp;&nbsp;
			</th>
        </tr>
        <tr valign="top" height="100%">
        	<td colspan="2" style="">
      	  		<iframe name="ifrm" src="help_parametros_liq_01.asp?concnro=<%= l_concnro%>" width="100%" height="100%"></iframe> 
			</td>
        </tr>
		<tr>
		    <td colspan="2" align="right" class="th2">
				<a class=sidebtnABM href="Javascript:Pasar_Valor();">Aceptar</a>
				<a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>
				<br> &nbsp;
			</td>
		</tr>
      </table>
</body>
</html>
