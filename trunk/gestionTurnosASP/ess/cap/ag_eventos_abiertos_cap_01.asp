<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo		: ag_eventos_abiertos_cap_01.asp
Descripcion	: Consulta de Eventos por empleados
Autor		: Lisandro Moro
Fecha		: 25/03/2004
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0
Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_mod
Dim l_filtro
Dim l_orden

l_ternro = l_ess_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY evecodext "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_EstiloTabla %>" rel="StyleSheet" type="text/css">
<title>Participación en cursos - Capacitación - RHPro &reg;</title>
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
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr nowrap>
    	<th align="center" nowrap>Cód. Ext</th>
		<th align="center">Descripción</th>
        <th align="center">Curso</th>
        <th align="center">Estado</th>
		<th align="center">Participación</th>
		<th align="center">Vacante</th>
		<th align="center" nowrap>Fecha Inicio</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT conf, estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, cap_curso.curdesabr, cap_evento.evefecini, evecanplaalu, evecanrealalu, cap_evento.evenro "
l_sql = l_sql & " from cap_candidato "
l_sql = l_sql & " inner join cap_evento on cap_evento.evenro = cap_candidato.evenro "
l_sql = l_sql & " INNER JOIN cap_curso ON cap_curso.curnro = cap_evento.curnro "
l_sql = l_sql & " INNER JOIN cap_estadoevento ON cap_estadoevento.estevenro = cap_evento.estevenro  "
l_sql = l_sql & " where ternro = " & l_ternro
l_sql = l_sql & " AND cap_evento.estevenro <> 6 "
l_sql = l_sql & " AND cap_evento.estevenro <> 4 "
l_sql = l_sql & " AND cap_evento.estevenro <> 7 "
l_sql = l_sql & " AND eveabierto = -1 "

if l_filtro <> "" then
  l_sql = l_sql & "AND " &l_filtro & " "
end if
l_sql = l_sql & l_orden

'response.write l_sql
'response.end

rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan=7>No hay Eventos para mostrar</td>
    </tr>
<%else
do until l_rs.eof
%>
	<tr onclick="Javascript:Seleccionar(this,<%= l_rs("evenro") %>);">
		<td align="left" width="10%" ><%= l_rs("evecodext")%></td>
     	<td align="left" width="38%" ><%= l_rs("evedesabr")%></td>
        <td align="left" width="20%" ><%= l_rs("curdesabr")%></td>
        <td align="left" width="10%" ><%= l_rs("estevedesabr")%></td>
		<td align="left" width="5%" ><% if l_rs("conf") = -1 then %>Participante<% Else %>Candidato<% End If %></td>
		<td align="center" width="5%" ><%= l_rs("evecanplaalu") - l_rs("evecanrealalu") %></td>
		<td align="left" width="20%" ><%= l_rs("evefecini")%></td>
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
