<%Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Codigo:;Descripcion:"
  l_Campos    = "reporte.repnro;reporte.repdesc"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "Codigo:;Descripcion:"
  l_CamposOr  = "reporte.repnro;reporte.repdesc"

dim l_otros
%>

<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Reportes</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>

function filtro(pag)
{
  abrirVentana('reporte_99.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('reporte_100.asp?pagina='+pag+'&lista=<%= l_orden %>&otros=<%=l_otros%>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
   	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Reportes</td>
<tr>
<td colspan="2" align="right" class="barra">
	<a class=sidebtnABM href="Javascript:abrirVentana('reporte_02.asp?Tipo=A','',545,130)">Alta</a>
	<a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'reporte_04.asp?repnro=' + document.ifrm.datos.cabnro.value)">Baja</a>
	<a class=sidebtnABM href="Javascript:abrirVentanaVerif('reporte_02.asp?Tipo=M&repnro=' + document.ifrm.datos.cabnro.value ,'',545,130)">Modifica</a>
	&nbsp;
	<a class=sidebtnABM href="Javascript:abrirVentanaVerif('columnas_reporte_00.asp?repnro=' + document.ifrm.datos.cabnro.value ,'',545,240)">Columnas</a>
	&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('reporte_01.asp');">Orden</a>
	<a class=sidebtnSHW href="Javascript:filtro('reporte_01.asp');">Filtro</a>
	&nbsp;
    <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>

<tr valign="top" height="100%">
   <td colspan="3" style="">
   <iframe name="ifrm" src="reporte_01.asp" width="100%" height="100%"></iframe> 
   </td>
</tr>

</table>
</body>
</html>
