<%Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%
'Archivo        : conf_mil_00.asp
'Descripcion    : Modulo que se encarga de admin. los servidores de mail
'Creador        : Lisandro Moro
'Fecha Creacion : 08/03/2005
'Modificacion   :

on error goto 0

dim l_rs
dim l_sql
dim l_tiptabdesc

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Código:;Descripción:;Origen:;Host:;Puerto:"
  l_Campos    = "cfgemailnro;cfgemaildesc;cfgemailfrom;cfgemailhost;cfgemailport"
  l_Tipos     = "N;T;T;T;N"

' Orden
  l_Orden     = "Código:;Descripción:;Origen:;Host:;Puerto:"
  l_CamposOr  = "cfgemailnro;cfgemaildesc;cfgemailfrom;cfgemailhost;cfgemailport"
  
%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Servicios de Mail</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function Alta(){
  abrirVentana('conf_mail_02.asp?Tipo=A','',320,220);
}

function Modificar(){
  if (document.ifrm.datos.cabnro.value == 0){
     alert('Debe seleccionar un registro.');
  }else{
     var param = '?Tipo=M&cfgemailnro=' + document.ifrm.datos.cabnro.value;  
     abrirVentana('conf_mail_02.asp' + param ,'',320,220);
  }
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)  
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
} 

function param(){
	return ('');
}

function salidaExcel(){
  abrirVentana('conf_mail_excel.asp?filtro='+ escape(document.ifrm.datos.filtro.value) +'&orden='+ document.ifrm.datos.orden.value ,'',300,300);
}    	   

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;">
<td align="left" class="barra" valign="top">Configuraci&oacute;n de Servicios de Mail</td>
<td align="right" class="barra">
	<a class=sidebtnABM href="Javascript:Alta();">Alta</a>
	<a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'conf_mail_04.asp?cfgemailnro='+document.ifrm.datos.cabnro.value)">Baja</a>
	<a class=sidebtnABM href="Javascript:Modificar();">Modifica</a>	
    &nbsp;&nbsp;&nbsp;
    <a class=sidebtnSHW href="Javascript:salidaExcel();">Excel</a>		  		  	
    &nbsp;&nbsp;&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('/turnos/seguridad/conf_mail_01.asp');">Orden</a>
	<a class=sidebtnSHW href="Javascript:filtro('/turnos/seguridad/conf_mail_01.asp');">Filtro</a>
    &nbsp;&nbsp;&nbsp;	
    <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>

</td>	  
</tr>
<tr valign="top">
   <td colspan="2" style="" height="100%">
      <iframe name="ifrm" src="conf_mail_01.asp" width="100%" height="100%"></iframe> 
   </td>
</tr>
</table>
</body>
</html> 
