<%Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : conf_x_empr_00.asp
Descripcion    : Modulo que se encarga de los abm de ConfPer.
Creador        : Scarpa D.
Fecha Creacion : 21/08/2003
Modificacion   :
   01/10/2003 - Scarpa D. - Agregado de los botones Filtro, Orden y Excel.
-----------------------------------------------------------------------------
-->
<%
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Código:;Descripción:;Valor:"
  l_Campos    = "confper.confnro;confper.confdesc;confper.confint"
  l_Tipos     = "N;T;N"

' Orden
  l_Orden     = "Código:;Descripción:;Valor:"
  l_CamposOr  = "confper.confnro;confper.confdesc;confper.confint"
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Empresas</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>

function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160);
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
} 

function salidaExcel(){
  abrirVentana('conf_x_empr_excel.asp?filtro='+ escape(document.ifrm.datos.filtro.value) +'&orden='+ document.ifrm.datos.orden.value ,'',300,300);
}    	   

function Alta(){
  abrirVentana('conf_x_empr_02.asp?Tipo=A','',330,170);
}

function Modificar(){
  if (document.ifrm.datos.cabnro.value == 0){
     alert('Debe seleccionar un registro.');
  }else{
     var param = '?Tipo=M&confnro=' + document.ifrm.datos.cabnro.value;  
     abrirVentana('conf_x_empr_02.asp' + param,'',330,170);
  }
}    	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
<td colspan="2" align="left" class="barra">Configuraci&oacute;n de Empresas</td>
<tr>
<td colspan="2" align="right" class="barra">
	<a class=sidebtnABM href="Javascript:Alta();">Alta</a>
	<a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'conf_x_empr_04.asp?confnro='+document.ifrm.datos.cabnro.value)">Baja</a>
	<a class=sidebtnABM href="Javascript:Modificar();">Modifica</a>	
    &nbsp;&nbsp;&nbsp;
    <a class=sidebtnSHW href="Javascript:salidaExcel();">Excel</a>		  		  	
    &nbsp;&nbsp;&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('conf_x_empr_01.asp');">Orden</a>
	<a class=sidebtnSHW href="Javascript:filtro('conf_x_empr_01.asp');">Filtro</a>
    &nbsp;&nbsp;&nbsp;	
    <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>	
</td>
</tr>
<tr valign="top">
   <td colspan="2" style="" height="100%">
      <iframe name="ifrm" src="conf_x_empr_01.asp" width="100%" height="100%"></iframe> 
   </td>
</tr>
</table>
</body>
</html>
