<%Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
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
  l_etiquetas = "Remitente:;Tipo:;Codigo:;Descripcion:;Fecha:"
  l_Campos    = "usrnombre;cysfirmas.cystipnro;cysfircodext;cysfirdes;cysfirfecaut;"
  l_Tipos     = "T;N;T;T;F"

' Orden
  l_Orden     = "Remitente:;Tipo:;Codigo:;Descripcion:;Fecha:"
  l_CamposOr  = "usrnombre;cysfirmas.cystipnro;cysfircodext;cysfirdes;cysfirfecaut;"


%>

<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Firmas pendientes</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script>

function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value,'',350,160)
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}


    	   
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name=datos>
	<input type="hidden" name="seleccion">
</form>
<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
<td align="left" class="barra">Firmas Pendientes</td>
<td align="right" class="barra">
	<a class=sidebtnABM href="Javascript:if (ifrm.jsSelRow != null) abrirVentana('Admin_Firmas_02.asp?Tipo=' + document.ifrm.jsSelRow.cells(7).innerText + '&descripcion=' + document.ifrm.jsSelRow.cells(3).innerText + '&codigo=' + document.ifrm.jsSelRow.cells(2).innerText,'',545,240); else alert('Debe seleccionar un registro');">Autorizar</a>
	<a class=sidebtnABM href="Javascript:alert('Opcion no implementada')">Depurar</a>
	<a class=sidebtnABM href="Javascript:if (ifrm.jsSelRow != null) abrirVentana('Admin_Firmas_04.asp?Tipo=' + document.ifrm.jsSelRow.cells(7).innerText + '&descripcion=' + document.ifrm.jsSelRow.cells(3).innerText + '&codigo=' + document.ifrm.jsSelRow.cells(2).innerText,'',545,240); else alert('Debe seleccionar un registro');">Detalle</a>
	&nbsp;&nbsp;&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('sup/admin_firmas_01.asp');">Orden</a>
	<a class=sidebtnSHW href="Javascript:filtro('admin_firmas_01.asp')">Filtro</a>
	&nbsp;&nbsp;&nbsp;
	<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
</td>
</tr>

<tr valign="top">
   <td colspan="2" style="">
   <iframe name="ifrm" src="Admin_Firmas_01.asp" width="100%" height="250""></iframe> 
   </td>
</tr>

</table>
</body>
</html>
