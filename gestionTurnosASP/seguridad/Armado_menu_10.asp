<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
' Variables
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Nombre:;Página:"
  l_Campos    = "bntnombre;btnpagina"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "Nombre:;Página:"
  l_CamposOr  = "bntnombre;btnpagina"

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Botones - Ticket</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
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
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td colspan="2" align="left" class="barra">Botones de <%= request("menuname")%></td>
          <td colspan="2" align="right" class="barra" nowrap>
		  <a class=sidebtnABM href="Javascript:abrirVentana('Armado_menu_12.asp?tipo=A&menuraiz=<%= request("menuraiz")%>&menuorder=<%= request("menuorder")%>','',500,190)">Alta</a>
		  <a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'Armado_menu_14.asp?menuraiz=<%= request("menuraiz")%>&menuorder=<%= request("menuorder")%>&nombre=' + escape(document.ifrm.datos.cabnro.value) + '&pagina=' + document.ifrm.datos.pagina.value)">Baja</a>
		  <a class=sidebtnABM href="Javascript:abrirVentanaVerif('Armado_menu_12.asp?tipo=M&menuraiz=<%= request("menuraiz")%>&menuorder=<%= request("menuorder")%>&cabnro=' + escape(document.ifrm.datos.cabnro.value) + '&pagina=' + document.ifrm.datos.pagina.value,'',500,190)">Modifica</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('Armado_menu_11.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('Armado_menu_11.asp');">Filtro</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="Armado_menu_11.asp?menuraiz=<%= request("menuraiz")%>&menuorder=<%= request("menuorder")%>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
      </table>
</body>
</html>
