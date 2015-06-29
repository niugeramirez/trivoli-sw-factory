<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: obrassociales_00.asp
'Descripción: ABM de Obras Sociales
'Autor : RAUL CHINESTRA
'Fecha: 19/04/2005
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
<title>Obras Sociales</title>
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
		abrirVentana("obrassociales_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra"></td>
          <td nowrap align="right" class="barra">
          <%'eugenio 29/06/2015, unificacion de iconos call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('obrassociales_02.asp?Tipo=A','',550,200);","Alta")%>
		  <a class="sidebtnABM" href="Javascript:abrirVentana('obrassociales_02.asp?Tipo=A','',550,200);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('obrassociales_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',550,200);","Modifica")%>
		  <a href="Javascript:abrirVentanaVerif('obrassociales_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',550,200);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'obrassociales_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  <a href="Javascript:eliminarRegistro(document.ifrm,'obrassociales_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>								  
		  &nbsp;&nbsp;
  		  <%'eugenio 29/06/2015, unificacion de iconos call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('listadeprecios_con_00.asp?id=' + document.ifrm.datos.cabnro.value,'',520,200);","Lista de Precios")%>
          <a href="Javascript:abrirVentanaVerif('listadeprecios_con_00.asp?id=' + document.ifrm.datos.cabnro.value,'',520,200);"><img src="/turnos/shared/images/Ecommerce-Price-Tag-icon_24.png" border="0" title="Lista de Precios"></a>								  

		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe name="ifrm" src="obrassociales_01.asp" width="100%" height="100%"></iframe> 
	      </td>
        </tr>
		
      </table>
</body>
</html>
