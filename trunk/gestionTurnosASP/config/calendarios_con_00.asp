<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
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
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Calendarios</title>
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
          <td align="left" class="barra">Calendarios</td>
          <td nowrap align="right" class="barra">
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value +'&Tipo=A','',520,200);","Alta")%>
		  &nbsp;
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'templatereservasdetalle_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  &nbsp;
          <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,200);","Modifica")%>
		  
		  <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('templatereservasdetalle_con_02.asp?idtemplatereserva='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,200);","Generar Calendario")%>
		  &nbsp;&nbsp;
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
      	  <iframe scrolling="Yes" name="ifrm" src="calendarios_con_01.asp?id=<%= l_id %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
  
</body>

</html>
