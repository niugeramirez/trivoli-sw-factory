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
  l_orden = " ORDER BY tercero.terape "
end if

l_evenro = request.querystring("evenro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Formadores asociados al Evento - Capacitación - RHPro &reg;</title>
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
<table>
    <tr>
        <th>Apellido</th>
        <th>Nombre</th>				
        <th>Interno</th>		
    </tr>
<%

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT tercero.ternro, tercero.terape , tercero.ternom, profinterno "
l_sql = l_sql & " FROM profesor "
l_sql = l_sql & " INNER JOIN cap_evento_profesor ON cap_evento_profesor.ternro = profesor.ternro  "
l_sql = l_sql & " INNER JOIN tercero ON tercero.ternro = cap_evento_profesor.ternro  "
l_sql = l_sql & " WHERE cap_evento_profesor.evenro = " & l_evenro 


if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if

l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Formadores</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr  onclick="Seleccionar(this,<%= l_rs("ternro")%> );" >
	        <td width="40%" align="left"><%= l_rs("terape")%></td>
	        <td width="40%" align="left"><%= l_rs("terape")%></td>
	        <td width="20%"  align="center" nowrap><% if  l_rs("profinterno") = - 1 then %>Si <% Else%>No<% End If %></td>

	    </tr>
	<%
		l_rs.MoveNext
	loop

  end if 
	

l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
