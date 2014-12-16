<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Gap registrados por competencia.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_excel_mat.asp
Descripción: Abm de Módulos - Excel.
Autor : Raul Chinestra
Fecha: 15/09/2003
Modificado: 16/08/2007 - Gustavo Ring - Se verifica que RR.HH haya aprobado
-->
<% 
Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_filtro
Dim l_orden

Dim l_saltear
Dim l_tienecomp
Dim l_empleg

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



'l_ternro = request("ternro")
l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY entnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
    <tr>
    	<th align="center" width="10%">Código</th>
        <th align="left">Descripción </th>
        <th align="center">Fecha</th>
        <th align="center"> % </th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_tienecomp = 0

' ************************************************* Manual  *********************************************

l_sql = "SELECT evafacnro , evafacdesabr, cap_capacita.fecha, cap_capacita.porcen "
l_sql = l_sql & " FROM cap_capacita "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_capacita.entnro "
l_sql = l_sql & " WHERE cap_capacita.origen1 = 5 " 'MANUAL 
l_sql = l_sql & " AND cap_capacita.origen2  = 3 " ' Competencias
l_sql = l_sql & " AND cap_capacita.idnro1  = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0

do until l_rs.eof
	l_tienecomp = -1
%>
	<tr >
		 <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="10%"><%= l_rs("fecha")%></td>
        <td align="center" width="20%"><%= l_rs("porcen") %></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close

' ************************************************* Eventos  *********************************************

l_sql = "SELECT evafacnro , evafacdesabr, cap_capacita.fecha, cap_capacita.porcen "
l_sql = l_sql & " FROM cap_capacita "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_capacita.entnro "
l_sql = l_sql & " WHERE cap_capacita.origen1 = 4 " 'EVENTOS
l_sql = l_sql & " AND cap_capacita.origen2  = 3 " ' Competencias
l_sql = l_sql & " AND cap_capacita.idnro1  = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0


do until l_rs.eof
	l_tienecomp = -1
%>
	<tr >
		 <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="10%"><%= l_rs("fecha")%></td>
        <td align="center" width="20%"><%= l_rs("porcen") %></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close

' ********************************************* Estudio Informal **************************************

l_sql = "SELECT evafacnro , evafacdesabr, cap_capacita.fecha, cap_capacita.porcen "
l_sql = l_sql & " FROM cap_estinformal "
l_sql = l_sql & " INNER JOIN cap_capacita ON cap_capacita.idnro2  = cap_estinformal.estinfnro "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_capacita.entnro "
l_sql = l_sql & " WHERE cap_capacita.origen1 = 3 " 'ESTUDIO INFORMAL
l_sql = l_sql & " AND cap_capacita.idnro1  = " & l_ternro 
l_sql =	l_sql & " AND cap_estinformal.estinfestrrhh = -1 "
l_sql = l_sql & " AND cap_capacita.origen2  = 3 " ' Competencias
l_sql = l_sql & " AND cap_estinformal.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0


do until l_rs.eof
	l_tienecomp = -1
%>
	<tr >
		 <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="10%"><%= l_rs("fecha")%></td>
        <td align="center" width="20%"><%= l_rs("porcen") %></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close

' ********************************************* Estudio Formal **************************************

l_sql = "SELECT evafacnro , evafacdesabr, cap_capacita.fecha, cap_capacita.porcen "
l_sql = l_sql & " FROM cap_estformal "
l_sql = l_sql & " INNER JOIN cap_capacita ON cap_capacita.idnro2  = cap_estformal.carredunro "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_capacita.entnro "
l_sql = l_sql & " WHERE cap_capacita.origen1 = 2 " 'ESTUDIO FORMAL
l_sql = l_sql & " AND cap_estformal.capestrrhh = -1" 'APROBADO POR RRHH
l_sql = l_sql & " AND cap_capacita.origen2  = 3 " ' Competencias
l_sql = l_sql & " AND cap_estformal.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0


do until l_rs.eof
	l_tienecomp = -1
%>
	<tr >
		 <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="70%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="15%"><%= l_rs("fecha")%></td>
        <td align="center" width="5%"><%= l_rs("porcen") %></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close

' ******************************************* Especializaciones ********************************

l_sql = "SELECT evafacnro , evafacdesabr, cap_capacita.fecha, cap_capacita.porcen "
l_sql = l_sql & " FROM especemp "
l_sql = l_sql & " INNER JOIN cap_capacita ON cap_capacita.idnro1 = especemp.eltananro AND cap_capacita.idnro2 = especemp.espnivnro "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_capacita.entnro "
l_sql = l_sql & " WHERE cap_capacita.origen1 = 1 " 'ESPECIALIZACIONES
l_sql = l_sql & " AND espestrrhh = -1 "
l_sql = l_sql & " AND cap_capacita.origen2  = 3 " ' Competencias
l_sql = l_sql & " AND especemp.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0


do until l_rs.eof
	l_tienecomp = -1
%>
	<tr >
		 <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="10%"><%= l_rs("fecha")%></td>
        <td align="center" width="20%"><%= l_rs("porcen") %></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close

' *******************************************************************************************************

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
