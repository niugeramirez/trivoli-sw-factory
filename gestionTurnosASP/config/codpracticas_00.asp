<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: codpracticas_00.asp
'Descripción: ABM de Codigos de Practicas
'Autor : Javier Posadas
'Fecha: 02/09/2019
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
  l_etiquetas = "Codigo:;Descripción:"
  l_Campos    = "balcod;baldes"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "Código:;Descripción:"
  l_CamposOr  = "balcod;baldes"

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Codigos Practicas</title>
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
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra"></td>
          <td nowrap align="right" class="barra">
          <a class="sidebtnABM" href="Javascript:abrirVentana('codpracticas_02.asp?Tipo=A','',550,200);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
		  &nbsp;          
          <a href="Javascript:abrirVentanaVerif('codpracticas_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',550,200);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
		  <a href="Javascript:eliminarRegistro(document.ifrm,'codpracticas_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>								  
		  &nbsp;&nbsp;
          </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe name="ifrm" src="codpracticas_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		
      </table>
</body>
</html>
