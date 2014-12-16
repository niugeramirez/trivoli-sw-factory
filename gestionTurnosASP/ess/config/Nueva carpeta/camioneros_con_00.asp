<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: camioneros_con_00.asp
'Descripci�n: Abm de camioneros
'Autor : Lisandro Moro
'Fecha: 15/02/2005

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
'
  l_etiquetas = "C�digo:;Apellido - Nombre:;Chasis:;Acoplado:"
  l_Campos    = "camcod;camdes;camcha;camaco"
  l_Tipos     = "N;T;T;T"

' Orden
  l_Orden     = "C�digo:;Apellido - Nombre:;Chasis:;Acoplado:;Habilitado:"
  l_CamposOr  = "camcod;camdes;camcha;camaco;camhab"

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%>Camioneros - Ticket</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag){
	abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("camioneros_con_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Camioneros</td>
          <td nowrap align="right" class="barra">
		  <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('camioneros_con_02.asp?tipo=C&cabnro=' + document.ifrm.datos.cabnro.value,'',500,400);","Consulta")%>
		  <% call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('terceros_documentos_con_00.asp?tipternro=2&cabnro=' + document.ifrm.datos.cabnro.value + '&descripcion=' + document.ifrm.datos.descripcion.value,'',600,360);","Documentos")%>
		  &nbsp;&nbsp;
		  <% call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('../../config/camioneros_con_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../config/camioneros_con_01.asp')">Filtro</a>
		  &nbsp;&nbsp;
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe name="ifrm" src="camioneros_con_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
	  <form name="datos">
	  	<input type="Hidden" name="cabnro">
	  </form>
</body>
</html>
