<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Especializaciones.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
----------------------------------------------------------------------------------------
Archivo: ag_especializaciones_cap_excel.asp
Descripcion: especializaciones
Autor: Lisandro Moro
Fecha: 29/03/2004
Modificación:  15/08/2007 - Gustavo Ring  - Se muestra el estado de aprobación de RR.HH	
------------------------------------------------------------------------------------------
-->

<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_orinro
Dim l_estnro
Dim l_saltear

Dim l_filtro
Dim l_orden

l_filtro 	= Request.QueryString("filtro")
l_orden  	= Request.QueryString("orden")
l_ternro 	= l_ess_ternro

if l_orden = "" then
	l_orden = "ORDER BY especemp.eltananro, eltanadesabr"
end if


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title> Estudios Informales - Capacitación - RHPro &reg;</title>
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
    	<th align="left">Código</th>
        <th align="left">Descripción Elemento</th>
        <th align="center">Código</th>
        <th align="center">Descripción Nivel</th>
   		<th align="center">RR.HH</th>
	</tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT especemp.eltananro, ternro, especemp.espnivnro, espmeses, espfecha, eltoana.eltanadesabr, espnivdesabr,espestrrhh"
l_sql = l_sql & " FROM especemp "
l_sql = l_sql & " inner join eltoana on eltoana.eltananro = especemp.eltananro"
l_sql = l_sql & " inner join espnivel on espnivel.espnivnro = especemp.espnivnro"
l_sql = l_sql & " WHERE ternro = " & l_ternro 

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " "& l_orden


rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan="5">No hay especializaciones registrados</td>
    </tr>
<%else
l_rs.MoveFirst
do until l_rs.eof
%>
	<tr onClick="Javascript:Seleccionar(this,<%= l_rs(0) %>)">
        <td align="left" width="10%"><%= l_rs(0) %></td>	
        <td align="left" width="40%"><%= l_rs(5) %></td>
        <td align="left" width="10%"><%= l_rs(2) %></td>
        <td align="left" width="40%"><%= l_rs(6) %></td>
		<%if not isnull(l_rs("espestrrhh")) then%> 
  			<%if l_rs("espestrrhh") = -1 then%> 
  				   <td nowrap align="center">Aceptado</td>			
			<%else%>
  				   <td nowrap align="center">Pendiente</td>						
			<%end if%>
		<%else%>
            <td nowrap align="center">Pendiente</td>						
		<%end if%>
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
</form>
</body>
</html>
