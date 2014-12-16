<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

'Archivo: perf_usr_seg_00.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005

' Variables
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Descripci&oacute;n:;Pol&iacute;tica de Cuenta:"
  l_Campos    = "perfnom;pol_desc"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "Descripci&oacute;n:;Pol&iacute;tica de Cuenta:"
  l_CamposOr  = "perfnom;pol_desc"

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Perfiles de Usuarios - Ticket</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value,'',350,160)
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	abrirVentana("perf_usr_seg_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%">
        <tr style="border-color :CadetBlue;">
          <td colspan="2" align="left" class="barra">Perfiles de Usuarios</td>
          <td colspan="2" align="right" class="barra">
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('perf_usr_seg_02.asp?Tipo=A','',520,140);","Alta")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'perf_usr_seg_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
          <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('perf_usr_seg_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,140);","Modifica")%>
  		  &nbsp;
          <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('../../seguridad/perf_usr_seg_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../seguridad/perf_usr_seg_01.asp');">Filtro</a>
		  &nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="perf_usr_seg_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
        <tr height="10">
          <td colspan="4"></td>
        </tr>
      </table>
</body>
</html>
