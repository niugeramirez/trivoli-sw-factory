<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: rhproappserv_seg_00.asp
'Descripción: Se encarga de ejecutar los procesos del sistema
'Autor : Lisandro Moro
'Fecha : 10/03/2005
'Modificado:


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
  l_etiquetas = "Desc. Abreviada:;Nivel:"
  l_Campos    = "cabpoldesabr;cabpolnivel"
  l_Tipos     = "T;N"

' Orden
  l_Orden     = "Descripción:;Nivel:"
  l_CamposOr  = "cabpoldesabr;cabpolnivel"

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>AppServer - Seguridad y Auditoría</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
  abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value,'',350,160)
}

function filtro(pag){
  abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
	<table border="0" cellpadding="0" cellspacing="0" height="100%">
		<tr style="border-color :CadetBlue;">
			<td colspan="2" align="left" class="barra">RHPro AppServer</td>
			<td colspan="2" align="right" class="barra">
				<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('rhproappsrv_seg_02.asp?Tipo=S&id=' + document.ifrm.datos.cabnro.value,'',545,200);","Start")%>
				<% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('rhproappsrv_seg_02.asp?Tipo=T&id=' + document.ifrm.datos.cabnro.value,'',545,200);","Stop")%>
				&nbsp;&nbsp;
				<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
			</td>
		</tr>
		<tr valign="top" height="100%">
			<td colspan="4" style="">
				<iframe name="ifrm" src="rhproappsrv_seg_01.asp" width="100%" height="100%"></iframe> 
			</td>
		</tr>
	</table>
</body>
</html>
