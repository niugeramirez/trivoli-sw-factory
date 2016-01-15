<%Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

'Archivo: pol_cuenta_seg_00.asp
'Descripción: ABM de Políticas de cuenta
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

' Variables
' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Descripci&oacute;n:;"
  l_Campos    = "pol_desc"
  l_Tipos     = "T"

' Orden
  l_Orden     = "Descripci&oacute;n:;"
  l_CamposOr  = "pol_desc"


%>
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Pol&iacute;ticas de Cuentas - Ticket</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
  abrirVentana('../Shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}


function filtro(pag){
  abrirVentana('../Shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}


function llamadaexcel(){ 
	abrirVentana("pol_cuenta_seg_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
    	<td align="left" class="barra">Pol&iacute;ticas de Cuentas</td>
        <td align="right" class="barra">
		  <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('pol_cuenta_seg_02.asp?Tipo=A','',720,330)","Alta") %>
		  <% call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'pol_cuenta_seg_04.asp?pol_nro=' + document.ifrm.datos.cabnro.value)","Baja") %>
	  	  <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('pol_cuenta_seg_02.asp?Tipo=M&pol_nro=' + document.ifrm.datos.cabnro.value,'',720,330)","Modifica") %>
  		  &nbsp;
          <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('../../seguridad/pol_cuenta_seg_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../seguridad/pol_cuenta_seg_01.asp')">Filtro</a>
		  &nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
    <tr valign="top" height="100%">
		<td colspan="2" style="">
      	  <iframe name="ifrm" src="pol_cuenta_seg_01.asp" width="100%" height="100%"></iframe> 
	    </td>
	</tr>
	<tr>
    	<td colspan="2" height="10"></td>
	</tr>
</table>
</body>
</html>
