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
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
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
 parent.actualizargap(cabnro); 

}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
		<th>Legajo</th>
        <th>Apellido</th>
        <th>Nombre</th>				
        <th>Hs. Totales</th>		
        <th>Hs. Asistió</th>				
        <th>  %  </th>				
    </tr>
<%

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT empleado.ternro, empleado.empleg , empleado.terape, empleado.ternom "
l_sql = l_sql & " FROM cap_candidato "
l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = cap_candidato.ternro  "
l_sql = l_sql & " WHERE cap_candidato.evenro = " & l_evenro & " AND cap_candidato.conf = -1 "


if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="6">No existen Participantes</td>
</tr>
<%else
    l_cantidademp = 0
	do until l_rs.eof
		
		' Calculo el numero total de minutos que Debe ir el Participante  
		l_sql2 = " SELECT cap_calendario.calnro, calhordes, calhorhas "
		l_sql2 = l_sql2 & " FROM cap_eventomodulo "
		l_sql2 = l_sql2 & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "
		l_sql2 = l_sql2 & " INNER JOIN cap_partcal ON cap_partcal.calnro = cap_calendario.calnro AND cap_partcal.ternro = " & l_rs("ternro")
		l_sql2 = l_sql2 & " WHERE cap_eventomodulo.evenro = " & l_evenro 
		rsOpen l_rs2, cn, l_sql2, 0 
		l_tot = 0
		l_can = 0
		do until l_rs2.eof
		
			l_tot = l_tot + datediff("n",cdate(mid(l_rs2("calhordes"),1,2)&":"& mid(l_rs2("calhordes"),3,2)),cdate(mid(l_rs2("calhorhas"),1,2)&":"& mid(l_rs2("calhorhas"),3,2)))

			' Calculo el numero total de minutos que Asistio el Empleado  
			l_sql3 = " SELECT asipre "
			l_sql3 = l_sql3 & " FROM cap_asistencia "
			l_sql3 = l_sql3 & " WHERE cap_asistencia.ternro = " & l_rs("ternro") & " AND cap_asistencia.calnro = " & l_rs2("calnro")
			rsOpen l_rs3, cn, l_sql3, 0 
			if not(l_rs3.eof) then
				if  l_rs3("asipre") = -1 then 
					l_can = l_can + datediff("n",cdate(mid(l_rs2("calhordes"),1,2)&":"& mid(l_rs2("calhordes"),3,2)),cdate(mid(l_rs2("calhorhas"),1,2)&":"& mid(l_rs2("calhorhas"),3,2)))							
					end if
			end if
			l_rs3.close
			l_rs2.MoveNext
		
		loop
		l_rs2.Close
		
		l_por = l_can * 100 / l_tot
		
		if l_por >= l_portot then
			l_cantidademp = 1

	%>
	    <tr  onclick="Seleccionar(this,<%= l_rs("ternro")%> );" >
	        <td width="10%" align="center"><%= l_rs("empleg")%></td>
	        <td width="30%" align="left"><%= l_rs("terape")%></td>
	        <td width="30%" nowrap><%= l_rs("ternom")%></td>
	        <td width="10%" nowrap align="center"><%= l_tot / 60 %></td>			
	        <td width="10%" nowrap align="center"><%= l_can / 60 %></td>			
	        <td width="10%" nowrap align="center"><%= l_por %></td>											
	    </tr>
	<%
		end if 
		l_rs.MoveNext
	loop
	if l_cantidademp = 0 then %>
	<tr> <td colspan="6">No existen Participantes</td></tr>	
<%  end if 
	
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
