<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<% 

'Archivo: novedades_empleado_liq_00.asp
'Descripción: abm novedades del empleado
'Autor: Fernando Favre
'Fecha: 10-10-03
'Modificado: 
'	17-11-03 FFavre Se agrego segundo apellido y nombre
'	25-11-03 FFavre Se agrando el tamaño de la llamada a ALTA y MODIFICACION
'   03-09-04 - Scarpa D. - Pasar como parametro el nenro
'	23-11-04 Alvaro Bayon - Cambio en el length de empleado
'	25-10-05 - Leticia A. - Adecuacion a Autogestion 


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
 l_etiquetas = "C&oacute;digo:;Concepto:;Par&aacute;metro:;Valor:;Vigencia Desde:;Vigencia Hasta:;Retroactivo Desde:;Retroactivo Hasta:;"
 l_Campos    = "concepto.conccod;concepto.concabr;tipopar.tpadabr;novemp.nevalor;novemp.nedesde;novemp.nehasta;periodo.pliqdesc;periodo.pliqdesc"
 l_Tipos     = "T;T;T;N;F;F;T;T"
 
' Orden
 l_Orden     = "C&oacute;digo:;Concepto:;Par&aacute;metro:;Valor:;Vigencia Desde:;Vigencia Hasta:;Retroactivo Desde:;Retroactivo Hasta:;"
 l_CamposOr  = "concepto.conccod;concepto.concabr;tipopar.tpadabr;novemp.nevalor;novemp.nedesde;novemp.nehasta;nepliqdesdedesc;nepliqhastadesc"
 
 Dim l_rs
 Dim l_sql
 
 Dim l_empleg
 
 l_empleg = request("empleg")
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Novedades por Empleado - Liquidaci&oacute;n de Haberes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>

function filtro(pag){
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag){
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
 	   
function param(){
	chequear= "empleg=<%= l_empleg%>";  //	chequear= "ternro=<%'= l_ternro%>";
	return chequear;
}
 	   
function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else{
	var par ="empleg=<%=l_empleg%>&orden="+document.ifrm.datos.orden.value+"&filtro="+escape(document.ifrm.datos.filtro.value);
		abrirVentana("novedades_empleado_liq_excel.asp?"+par,'execl',250,150);
	}
}


function Tecla(num){
	if (num==13) {
		return false;
  }
  return num;
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<!--input type="hidden" name="seleccion"-->

<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
    	<th align="left">Novedades por Empleado</th>
        <th nowrap style= " text-align: right;">
		<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('novedades_empleado_liq_02.asp?tipo=A&empleg=" & l_empleg &"','',650,250);","Alta")%>
		<% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'novedades_empleado_liq_04.asp?empleg="& l_empleg &"&nenro='+document.ifrm.datos.cabnro.value);","Baja")%>
		<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('novedades_empleado_liq_02.asp?tipo=M&nenro='+document.ifrm.datos.cabnro.value+'&empleg="& l_empleg &"&concnro='+document.ifrm.datos.concnro.value+'&tpanro='+document.ifrm.datos.tpanro.value,'',650,250);", "Modifica")%>
		&nbsp;&nbsp;&nbsp;
		<!-- a class=sidebtnSHW href="Javascript:abrirVentana('conceptos_mopre_liq_00.asp?obj=document.datos.seleccion','',500,420);">Conceptos</a>  -->
		<% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();", "Excel")%>
		&nbsp;&nbsp;&nbsp;
		<a class=sidebtnSHW href="Javascript:orden('/ess/ess/liq/novedades_empleado_liq_01.asp');">Orden</a>
		<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/liq/novedades_empleado_liq_01.asp')">Filtro</a>
		&nbsp;&nbsp;
		</th>
	</tr>
	<tr valign="top" height="100%">
		<td colspan="2" style="">
      		<iframe name="ifrm" src="novedades_empleado_liq_01.asp?empleg=<%=l_empleg %>" width="100%" height="100%"></iframe> 
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20"></td>
	</tr>
</table>
</form>
</body>
</html>
