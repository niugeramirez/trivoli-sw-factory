<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/ess/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: agencia_con_00.asp
'Descripción: ABM de Agencias
'Autor : Raul Chinestra
'Fecha: 20/08/2008

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
  
  Dim l_id
  
  l_id  = request("id")

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Detalle Modelo de Turnos</title>
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
<form name="datos" method="post">
<input type="hidden" name="id" value="<%= l_id %>">
</form>	
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Modelos de Turnos<%= l_id %></td>
          <td nowrap align="right" class="barra">
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value +'&Tipo=A','',520,200);","Alta")%>
		  <a class="sidebtnABM" href="Javascript:abrirVentana('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value +'&Tipo=A','',550,200);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,200);","Modifica")%>
		  <a href="Javascript:abrirVentanaVerif('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',550,200);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'templatereservasdetalle_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  <a href="Javascript:eliminarRegistro(document.ifrm,'templatereservasdetalle_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>								  		  
		  &nbsp;&nbsp;
          		 
		  &nbsp;&nbsp;
		  
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="Yes" name="ifrm" src="templatereservasdetalle_con_01.asp?id=<%= l_id %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
  
</body>

</html>
