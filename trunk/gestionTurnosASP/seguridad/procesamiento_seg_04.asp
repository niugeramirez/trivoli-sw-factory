<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: procesamiento_seg_04.asp
'Descripción: Se encarga de mostrar los procesos del sistema
'Autor : Lisandro Moro
'Fecha : 10/03/2005
'Modificado:

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Filtro
  l_etiquetas = "Fecha:;Proceso:;Usuario:;Estado:"
  l_Campos    = "bprcfecha;btprcdesabr;iduser;bprcestado"
  l_Tipos     = "F;T;T;T"

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Servidor de Aplicaciones</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
	if (jsFecha == null)
		txt.value = ''
	else
		txt.value = jsFecha;
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>','',250,160);
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr>
	    <td align="left" class="barra">Servidor de Aplicaciones</td>
	    <td nowrap align="right" class="barra" valign="middle">
		<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('procesamiento_seg_10.asp?bpronro=' + document.ifrm.datos.cabnro.value,'',500,340);","Control")%>
		<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('rhproappsrv_seg_00.asp','',470,340);","Start/Stop")%>
		&nbsp;&nbsp;&nbsp;
		<!--<a class=sidebtnSHW href="javascript:abrirVentanaVerif('procesamiento_seg_06.asp?bpronro=' + document.ifrm.datos.cabnro.value,'',400,365)">Empleados</a>-->
		<!--<a class=sidebtnSHW href="javascript:abrirVentanaVerif('procesamiento_seg_08.asp?bpronro=' + document.ifrm.datos.cabnro.value,'',400,365)">Procesos</a>-->
		&nbsp;&nbsp;&nbsp;
		<a class=sidebtnSHW href="javascript:Javascript:filtro('../../seguridad/procmon.asp')">Filtro</a>
		&nbsp;&nbsp;&nbsp;
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr height="100%">
		<td colspan="2">
	  				<iframe name="ifrm" src="procmon.asp" width="100%" height="100%"></iframe> 
		</td>
	</tr>
</table>
</body>
</html>
