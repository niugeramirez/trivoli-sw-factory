<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: asistencias_control_cap_01.asp
Descripción: Control de Asistencias
Autor : Raul CHinestra
Fecha: 13/01/2004
-->
<% 

Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_sql
Dim l_sql2
Dim l_sql3
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_evenro
Dim l_tot
Dim l_can
Dim l_por
Dim l_portot
Dim l_cantidademp

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY empleado.empleg "
end if

l_evenro = request.querystring("evenro")


Set l_rs  = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT eveporasi "
l_sql = l_sql & " FROM cap_evento "
l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
rsOpen l_rs, cn, l_sql, 0 
if not(l_rs.eof) then
	l_portot = l_rs("eveporasi")
end if
l_rs.close


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray4.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Modulos asociados al Evento - Capacitación - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila; 
 //parent.actualizargap(cabnro); 

}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table height="100%" width="100%" cellpadding="0" cellspacing="0">

   <tr valign="top" height="50%">
         <td colspan="2" style="" align="center">
     	  <iframe  frameborder="0" scrolling="Yes" name="ifrm" frameborder="0" src="ag_cerrar_asi_cap_00.asp?evenro= <%= l_evenro %> " width="100%" height="100%"></iframe> 		  		       	
      </td>
   </tr>
    <tr valign="top" height="50%">
         <td colspan="2" style="" align="center">
     	  <iframe frameborder="0" scrolling="Yes" name="ifrm1" frameborder="0" src="ag_cerrar_asi_cap_01.asp?evenro= <%= l_evenro %> " width="100%" height="100%"></iframe> 
      </td>
    </tr>		
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
<script>
window.document.body.scroll = "no";
</script>
