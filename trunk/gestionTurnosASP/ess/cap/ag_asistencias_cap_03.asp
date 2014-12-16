<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: asistencias_cap_03.asp
Descripción: Abm de Asistencias asociadas al Evento
Autor : Raul CHinestra
Fecha: 11/12/2003
-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_evenro
Dim l_todos
Dim l_eveorigen
Dim l_eveforeva

l_filtro = request("filtro")
l_orden  = request("orden")

l_evenro = request.querystring("evenro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT eveorigen, eveforeva"
l_sql = l_sql & " FROM cap_evento "
l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	l_eveforeva = l_rs("eveforeva")	
	if isnull(l_rs("eveorigen"))   then 
		l_eveorigen = 0
	else 
		l_eveorigen = l_rs("eveorigen")	
	end if
end if 

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Asistencias al Evento - Capacitación - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_sel_multiple.js"></script>
<script>
var jsSelRow = null;
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
     	<th>Apellido</th>
     	<th>Nombre</th>
        <th>Fecha</th>
        <th>Día</th>
		<th>H.Desde</th>
		<th>H.Hasta</th>
		<th>Lugar</th>
		<th>Presente</th>
		<th>Avisó</th>		
    </tr>
<%


l_rs.close

l_sql =  " select cap_asistencia.empleado, tercero.terape, tercero.ternom, cap_asistencia.calnro,cap_calendario.calfecha, cap_calendario.caldia, cap_asistencia.asievehorini, cap_asistencia.asievehorfin, cap_lugar.lugdesabr, cap_asistencia.asipre, cap_asistencia.asiavi"
l_sql = l_sql & " from cap_asistencia "
l_sql = l_sql & " INNER JOIN cap_calendario ON cap_calendario.calnro = cap_asistencia.calnro "
l_sql = l_sql & " INNER JOIN tercero ON tercero.ternro = cap_asistencia.empleado "
l_sql = l_sql & " INNER JOIN cap_lugar ON cap_lugar.lugnro = cap_calendario.lugnro "
l_sql = l_sql & " where cap_asistencia.empleado IN "
l_sql = l_sql & " (SELECT cap_candidato.ternro "
l_sql = l_sql & " FROM cap_candidato "
l_sql = l_sql & " WHERE ((cap_candidato.evenro = " & l_evenro & " AND cap_candidato.evento_origen = 1)"
l_sql = l_sql & " OR (cap_candidato.evenro = " & l_eveorigen & " AND cap_candidato.evento_origen = 2 ))"
l_sql = l_sql & " and (cap_candidato.conf = -1) and cap_asistencia.calnro IN "
l_sql = l_sql & " (SELECT cap_calendario.calnro "
l_sql = l_sql & " FROM cap_evento "
l_sql = l_sql & " INNER JOIN cap_calendario ON cap_evento.evenro = cap_calendario.evenro "
l_sql = l_sql & " WHERE ((cap_evento.evenro = " & l_evenro & " AND cap_evento.eveforeva = 1 )"
l_sql = l_sql & " OR (cap_evento.eveorigen = " & l_eveorigen & " AND cap_evento.eveforeva = 2 ))  ))"

if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 


if l_rs.eof then%>
<tr>
	 <td colspan="10">No existen Asistencias</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr onclick="Javascript:Seleccionar(this,'<%= l_rs("ternro") & "-" & l_rs("calnro") %>');">
			<td width="20%" ><%= l_rs("terape")%></td>
			<td width="20%" nowrap><%= l_rs("ternom")%></td>			
			<td width="10%" align="right"><%= l_rs("calfecha")%></td>
			<td width="10%" nowrap><%= l_rs("caldia")%></td>			
	        <td width="5%" nowrap><%= mid(l_rs("asievehorini"),1,2)&":"& mid(l_rs("asievehorini"),3,2)%></td>
			<td width="5%" nowrap><%= mid(l_rs("asievehorfin"),1,2)&":"& mid(l_rs("asievehorfin"),3,2)%></td>
			<td width="20%" nowrap><%= l_rs("lugdesabr")%></td>
			<td width="5%" nowrap><% if l_rs("asipre") = -1 then %>Si <% else %>No<% end if %></td>			
			<td width="5%" nowrap><% if l_rs("asiavi") = -1 then %>Si <% else %>No<% end if %></td>			
	    </tr>		
	<%
 	    if l_todos="" then
           l_todos = l_rs("ternro") & "-" & l_rs("calnro") 
   	    else
           l_todos = l_todos & "," & l_rs("ternro") & "-" & l_rs("calnro") 
  	    end if
		
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
<input type="hidden" name="listanro" value="">
<input type="hidden" name="listatodos" value="<%= l_todos%>">
</form>

<script>
  setearObjDatos(document.datos.listanro, document.datos.listatodos);
</script>
</body>
</html>
