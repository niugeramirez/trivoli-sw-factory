<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!-- ----------------------------------------------------------------------------
Archivo			: empleados_reportan_mss_00.asp
Descripción		: Consulta que muestra los empleados que reportan al empleado logueado.
Autor			: Fernando Favre
Fecha			: 30-05-2005
Modificado		:
-----------------------------------------------------------------------------  -->
<%
on error goto 0
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
 Dim l_Orden      ' Son las etiquetas que aparecen en el orden
 Dim l_CamposOr   ' Son los campos para el orden
 
' Filtro
 l_etiquetas = "Legajo:;Apellido:"
 l_Campos    = "empleg;terape"
 l_Tipos     = "N;T"
 
' Orden
 l_Orden     = "Legajo:;Apellido:"
 l_CamposOr  = "empleg;terape,terape2,ternom,ternom2"
 
%>
<html>
<head>
<link href="../<%= c_estilo%>" rel="StyleSheet" type="text/css">
<title>Empleados que reportan - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
function orden(pag){
	abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr>
          <th align="left">Empleados que reportan</th>
          <!--td nowrap align="center" class="barra">
			  &nbsp;
	          <%' call MostrarBoton ("sidebtnSHW", "Javascript:Recibo();","Ver Recibo")%>
			  &nbsp;
		  </td-->
          <th nowrap align="right">
			  <a class=sidebtnSHW href="Javascript:orden('../../mss/empleados_reportan_mss_01.asp');">Orden</a>
			  <a class=sidebtnSHW href="Javascript:filtro('../../mss/empleados_reportan_mss_01.asp')">Filtro</a>
		  </th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="">
      	  <iframe name="ifrm" src="empleados_reportan_mss_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		<tr>
			<td colspan="2" height="20"></td>
		</tr>
      </table>
</body>
</html>
