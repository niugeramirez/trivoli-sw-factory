<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=eventos_competencia.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo		: con_eventos_en_competencia_cap_excel.asp
Descripcion	: Consulta de Eventos por competencia
Autor		: Juan Manuel Hoffman
Fecha		: 19/03/2004
-----------------------------------------------------------------------------
-->

<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_compe
Dim l_filtro
Dim l_orden
Dim l_empleg

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_compe  = request("competencia")

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY evecodext "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Eventos por Empleado - Capacitación - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
   <tr>
		<th colspan="6" align="center">Ewventos asociados a la Competencia</th>
	</tr>
    <tr>
    	<th align="center">Cód. Ext</th>
		<th align="center">Descripción</th>
        <th align="center">Curso</th>
        <th align="center">Estado</th>
		<th align="center">Vacante</th>
		<th align="center">Fecha Inicio</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, cap_curso.curdesabr, cap_evento.evefecini, evecanplaalu, evecanrealalu "
l_sql = l_sql & " FROM cap_evento  "
l_sql = l_sql & " INNER JOIN cap_dictado ON cap_dictado.evenro = cap_evento.evenro AND cap_dictado.origen = 3 AND cap_dictado.entnro =" & l_compe
l_sql = l_sql & " INNER JOIN cap_curso ON cap_curso.curnro = cap_evento.curnro "
l_sql = l_sql & " INNER JOIN cap_estadoevento ON cap_estadoevento.estevenro = cap_evento.estevenro "
l_sql = l_sql & " WHERE cap_evento.estevenro <> 6 "

if l_filtro <> "" then
  l_sql = l_sql & "AND " &l_filtro & " "
end if
l_sql = l_sql & l_orden

rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan=4>No hay Eventos registrados para esa Competencia</td>
    </tr>
<%else
do until l_rs.eof
%>
	<tr>
		<td align="left" width="25%"><%= l_rs("evecodext")%></td>
     	<td align="left" width="30%"><%= l_rs("evedesabr")%></td>
        <td align="left" width="25%"><%= l_rs("curdesabr")%></td>
        <td align="left" width="20%"><%= l_rs("estevedesabr")%></td>
		<td align="center" width="20%"><%= l_rs("evecanplaalu") - l_rs("evecanrealalu") %></td>
		<td align="left" width="20%"><%= l_rs("evefecini")%></td>
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
</body>
</html>
