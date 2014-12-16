<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
' Variables
on error goto 0

Dim l_otros

' Filtro
Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
Dim l_Orden      ' Son las etiquetas que aparecen en el orden
Dim l_CamposOr   ' Son los campos para el orden

' Filtro
l_etiquetas = "Conf.Auditoria:;Descripcion:"
l_Campos    = "confaud.caudnro;confaud.cauddes"
l_Tipos     = "N;T"

' Orden
l_Orden     = "Conf.Auditoria:;Descripcion:;Activo:"
l_CamposOr  = "confaud.caudnro;confaud.cauddes;confaud.caudact"
%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Auditor&iacute;a</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function filtro(pag){
  abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag){
  abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&otros=<%=l_otros%>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<input type="hidden" name="seleccion">
</form>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr>
		<td align="left" class="barra">Configuraci&oacute;n de Auditor&iacute;a</td>
		<td align="right" class="barra" nowrap>
			<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('configuracion_auditoria_02.asp?Tipo=A','',500,140);","Alta")%>
			<% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'configuracion_auditoria_04.asp?caudnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
			<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('configuracion_auditoria_02.asp?Tipo=M&caudnro=' + document.ifrm.datos.cabnro.value ,'',500,140);","Modifica")%>
			&nbsp;&nbsp;&nbsp;
			<a class=sidebtnSHW href="Javascript:orden('../../seguridad/configuracion_auditoria_01.asp');">Orden</a>
			<a class=sidebtnSHW href="Javascript:filtro('../../seguridad/configuracion_auditoria_01.asp');">Filtro</a>
			&nbsp;&nbsp;
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
	   <td colspan="2" height="100%">
		   <iframe name="ifrm" src="configuracion_auditoria_01.asp" width="100%" height="100%" ></iframe> 
	   </td>
	</tr>
</table>
</body>
</html>
