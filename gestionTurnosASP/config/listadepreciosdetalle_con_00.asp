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
  Dim l_rs
  Dim l_sql
  
  Dim l_obra_social
  Dim l_titulo
  
  l_id  = request("id")
  
  
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT * "
  l_sql = l_sql & " FROM listaprecioscabecera "
  l_sql = l_sql & " INNER JOIN obrassociales ON obrassociales.id = listaprecioscabecera.idobrasocial "
  l_sql = l_sql & " WHERE listaprecioscabecera.id = " & l_id
  rsOpen l_rs, cn, l_sql, 0 
  if not l_rs.eof then
  	l_obra_social = l_rs("descripcion")
	l_titulo = l_rs("titulo")
  else
  	l_obra_social = ""
	l_titulo = "" 
  	
  end if
  

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Detalle Lista de Precios</title>
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
          <td align="left" class="barra">Detalle Lista de Precios</td>
          <td nowrap align="right" class="barra">
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('listadepreciosdetalle_con_02.asp?idcab='+document.datos.id.value +'&Tipo=A','',520,200);","Alta")%>
		  <a class="sidebtnABM" href="Javascript:abrirVentana('listadepreciosdetalle_con_02.asp?idcab='+document.datos.id.value +'&Tipo=A','',520,200);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta">
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('listadepreciosdetalle_con_02.asp?idcab='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,100);","Modifica")%>
		  <a href="Javascript:abrirVentanaVerif('listadepreciosdetalle_con_02.asp?idcab='+document.datos.id.value + '&Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',520,100);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
          <%'eugenio 29/06/2015, unificacion de iconos  call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'listadepreciosdetalle_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  <a href="Javascript:eliminarRegistro(document.ifrm,'listadepreciosdetalle_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>	
		  &nbsp;&nbsp;
          
		  &nbsp;&nbsp;
		  
		  </td>
        </tr>
		<tr>
			<td align="right"><b>Obra Social: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_obra_social %>" readonly   class="deshabinp" >
			</td>
		</tr>		
		<tr>
			<td align="right"><b>Titulo: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_titulo %>" readonly   class="deshabinp">
			</td>
		</tr>			
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="Yes" name="ifrm" src="listadepreciosdetalle_con_01.asp?idcab=<%= l_id %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>
  
</body>

</html>
