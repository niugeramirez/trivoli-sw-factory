<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_00.asp
Descripción: Abm de Solicitudes
Autor : Raul Chinestra
Fecha: 30/03/2004

-->
<% 

on error goto 0

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
  l_etiquetas = "C&oacute;digo:;Descripción:"
  l_Campos    = "solnro;soldesabr"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "C&oacute;digo:;Descripción:"
  l_CamposOr  = "solnro;soldesabr"

Dim l_ternro
l_ternro  = request("ternro")


%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Solicitud de Eventos - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("contenidos_cap_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

</script>
</head>

<form name="datos">
<input type=hidden name=ternro value="<%= l_ternro %>">
</form>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <th align="left">Requerimiento de eventos que no estén en la oferta interna</th>
          <th nowrap align="right">
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('ag_solicitud_eventos_cap_02.asp?Tipo=A&ternro='+ document.datos.ternro.value,'',550,220);","Alta")%>
		  <a class=sidebtnABM href="Javascript:abrirVentana('ag_solicitud_eventos_cap_02.asp?Tipo=A&ternro='+ document.datos.ternro.value,'',550,200);">Alta</a>
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'ag_solicitud_eventos_cap_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  <a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'ag_solicitud_eventos_cap_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);">Baja</a>
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('ag_solicitud_eventos_cap_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value + '&ternro='+document.datos.ternro.value,'',550,220);","Modifica")%>
		  <a class=sidebtnABM href="Javascript:abrirVentanaVerif('ag_solicitud_eventos_cap_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value + '&ternro='+document.datos.ternro.value,'',550,200);">Modifica</a>
		  &nbsp;&nbsp;
  		  <!--
          <% 'call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  <a class=sidebtnABM href="Javascript:llamadaexcel();">Excel</a>
		  
		  &nbsp;&nbsp;&nbsp;

		  <a class=sidebtnSHW href="Javascript:orden('ag_solicitud_eventos_cap_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('ag_solicitud_eventos_cap_01.asp')">Filtro</a>
		  		  
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<% ' Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  -->
		  </th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe frameborder="0" name="ifrm"  scrolling="Yes" src="ag_solicitud_eventos_cap_01.asp?ternro=<%= l_ternro %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		<tr>
			<td colspan="2" height="20">
			</td>
		</tr>
      </table>
</body>
<script>
window.document.body.scroll = "no";
</script>
</html>
