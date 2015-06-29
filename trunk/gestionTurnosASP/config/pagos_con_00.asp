<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/ess/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
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
  l_etiquetas = "Descripción:;"
  l_Campos    = "agedes"
  l_Tipos     = "T;"

' Orden
  l_Orden     = "Descripción:;"
  l_CamposOr  = "agedes"

  Dim l_idpracticarealizada
  
  l_idpracticarealizada = request("cabnro")
  

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Detalle de Pagos</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
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
		abrirVentana("agencias_con_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">
		  
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('pagos_con_02.asp?Tipo=A&idpracticarealizada=" & l_idpracticarealizada & "','',520,350);","Alta")%>
		  <a class="sidebtnABM" href="Javascript:abrirVentana('pagos_con_02.asp?Tipo=A&idpracticarealizada=<%=l_idpracticarealizada%>','',520,350);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('pagos_con_02.asp?Tipo=M&idpracticarealizada=" & l_idpracticarealizada & "&cabnro=' + document.ifrm.datos.cabnro.value,'',520,350);","Modifica")%>
		  <a href="Javascript:parent.abrirVentanaVerif('pagos_con_02.asp?Tipo=M&idpracticarealizada=<%=l_idpracticarealizada%>&cabnro=' + document.ifrm.datos.cabnro.value,'',520,350);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'pagos_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  <a href="Javascript:eliminarRegistro(document.ifrm,'pagos_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>								  
		  &nbsp;&nbsp;
		  &nbsp;
          <%' call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  <!-- <a class=sidebtnSHW href="Javascript:orden('../../config/agencias_con_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../config/agencias_con_01.asp')">Filtro</a>
		   -->
		  &nbsp;&nbsp;
		  <!--
		  <a class=sidebtnHLP href="Javascript:ayuda('<%'= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		  -->
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="Yes" name="ifrm" src="pagos_con_01.asp?idpracticarealizada=<%= l_idpracticarealizada  %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
</body>
</html>
