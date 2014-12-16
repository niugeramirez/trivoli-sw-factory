<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<% 
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
  l_etiquetas = "N&uacute;mero:;Nombre:"
  l_Campos    = "mulnro;mulnom"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "N&uacute;mero:;Nombre:"
  l_CamposOr  = "mulnro;mulnom"

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Configuración multiempresa - Ticket</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('multiempresa_gti_100.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
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
          <td colspan="2" align="left" class="barra">Configuración Multiempresa</td>
          <td colspan="2" align="right" class="barra">
		  <a class=sidebtnABM href="Javascript:abrirVentana('multiempresa_gti_02.asp?Tipo=A','',400,100)">Alta</a>
		  <a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'multiempresa_gti_04.asp?mulnro=' + document.ifrm.datos.cabnro.value)">Baja</a>
		  <a class=sidebtnABM href="Javascript:abrirVentanaVerif('multiempresa_gti_02.asp?Tipo=M&mulnro=' + document.ifrm.datos.cabnro.value,'',400,100)">Modifica</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('multiempresa_gti_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('multiempresa_gti_01.asp')">Filtro</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="multiempresa_gti_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
      </table>
</body>
</html>
