<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: tablero_eventos_cap_01.asp
Descripción: eventos en que participo un Empleado
Autor : Raul Chinestra
Fecha: 16/01/2004
-->
<% 
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_ternro

dim leg
leg = Session("empleg")
if leg = "" then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	Response.End
end if
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro FROM empleado WHERE empleado.empleg =" & leg
l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "NO SE HA SELECCIONADO UN EMPLEADO<BR>"
	response.end
else 
  l_ternro = l_rs("ternro")
end if
l_rs.close
l_empleg     = leg



%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Eventos - Capacitación - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro){
	if (jsSelRow != null) {
		Deseleccionar(jsSelRow);
 	};
 document.datos.cabnro.value = cabnro;
 parent.document.datos.evenro.value = cabnro;
 parent.active();
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo Externo</th>
		<th>Descripción</th>
        <th>Curso</th>		
		<th>Estado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "select evecodext, evedesabr, curdesabr, cap_evento.evenro, estevedesabr "
l_sql = l_sql & " from cap_evento "
l_sql = l_sql & "  inner join cap_candidato on cap_candidato.evenro = cap_evento.evenro "
l_sql = l_sql & "         and cap_candidato.ternro = " & l_ternro
'l_sql = l_sql & "         and cap_candidato.conf = -1"
l_sql = l_sql & " inner join cap_curso on cap_curso.curnro = cap_evento.curnro	"
l_sql = l_sql & " inner join cap_estadoevento on cap_estadoevento.estevenro = cap_evento.estevenro	"
'l_sql = l_sql & " where cap_evento.estevenro = 6 "
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
		<tr>
			 <td colspan="4">No asistió a Eventos</td>
		</tr>
<%else
	do until l_rs.eof	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evenro")%>);">
	        <td width="15%" align="left"><%= l_rs("evecodext")%></td>
			<td width="15%" align="left" nowrap><%= l_rs("evedesabr")%></td>
	        <td width="40%" nowrap><%= l_rs("curdesabr")%></td>
			<td width="15%" nowrap><%= l_rs("estevedesabr") %></td>
	    </tr>
	<% l_rs.MoveNext
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
