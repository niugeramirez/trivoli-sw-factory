<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'Modificado : Gustavo Ring  - 16/08/2007 - Se agrego el estado de aprobación de RR.HH
'Martin Ferraro - 30/08/2007 - Correccion de tipos en ORACLE

on error goto 0
Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_orinro
Dim l_estnro
Dim l_saltear
Dim l_empleg

Dim l_filtro
Dim l_orden

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY estinfnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla %>" rel="StyleSheet" type="text/css">
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
        <th align="left">Descripción Abreviada</th>
        <th align="center">Tipo de Curso</th>
        <th align="center">Institución</th>
		<th align="center">RR.HH</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estinfnro, estinfdesabr, tipcurdesabr, instdes, estinfestrrhh"
l_sql = l_sql & " FROM cap_estinformal "
l_sql = l_sql & " INNER JOIN cap_tipocurso ON cap_tipocurso.tipcurnro = cap_estinformal.tipcurnro "
l_sql = l_sql & " INNER JOIN institucion ON institucion.instnro = cap_estinformal.instnro "
l_sql = l_sql & " WHERE ternro = " & l_ternro 
if l_filtro <> "" then
  l_sql = l_sql & "AND " &l_filtro & " "
end if
l_sql = l_sql & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan="5">No hay Estudios Informales registrados</td>
    </tr>
<%else
do until l_rs.eof
%>
	<tr onDblClick="Javascript:parent.abrirVentana('ag_estudios_informales_cap_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',620,280)" onClick="Javascript:Seleccionar(this,<%= l_rs("estinfnro")%>)">
        <td align="left" width="10%"><%= l_rs("estinfnro")%></td>	
        <td align="left" width="30%"><%= l_rs("estinfdesabr")%></td>
        <td align="left" width="30%"><%= l_rs("tipcurdesabr")%></td>
        <td align="left" width="20%"><%= l_rs("instdes") %></td>
		<%if not isnull(l_rs("estinfestrrhh")) then%> 
  			<%if clng(l_rs("estinfestrrhh")) = -1 then%> 
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
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
