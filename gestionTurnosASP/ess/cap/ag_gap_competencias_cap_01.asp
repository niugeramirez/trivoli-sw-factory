<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: gap_competencias_cap_01.asp
Descripcion: Gap por Competencias
Autor: Lisandro Moro
Fecha: 12/02/2004
Modificado: Raul Chinestra - 26/02/2004 - Se agrego el campo de Porcentaje
-----------------------------------------------------------------------------
-->
<% 
Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_orinro
Dim l_estnro
Dim l_saltear

'l_ternro = request("ternro")
dim leg
Dim l_empleg

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



l_estnro = request.querystring("estado")
if l_estnro = "2" then
	l_saltear= "Si"
else 	
    l_saltear= "No"
end if


l_orinro = request.querystring("origen")
if l_orinro = "" then
	l_orinro= "0"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gap Registrados por Competencias - Capacitación - RHPro &reg;</title>
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
    	<th align="center">Código</th>
        <th align="left">Competencia</th>
        <th align="center">Fecha</th>
        <th align="center">Estado</th>
        <th align="center"> % </th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evafacnro,  falorigen, evafacdesabr, falidnro, falfecha, falpendiente, falporcen"'falnro,
l_sql = l_sql & " FROM cap_falencia "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_falencia.modnro "
l_sql = l_sql & " WHERE ternro = " & l_ternro 
l_sql = l_sql & " AND falorigen = 7 "
'l_sql = l_sql & " AND (falorigen = " & l_orinro & " OR 0 = " & l_orinro & ")"
if l_saltear = "No" then 
	l_sql = l_sql & " AND falpendiente = " & l_estnro
end if 
l_sql = l_sql & " ORDER BY falorigen" ', falnro

rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan=5>No hay Gap registrados por Competencias</td>
    </tr>
<%else
do until l_rs.eof
%>
	<tr ><!--ondblclick="Javascript:parent.abrirVentana('gap_competencias_cap_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,180)" onclick="Javascript:Seleccionar(this,<%'= l_rs("falnro")%>)"-->
        <td align="center" width="10%"><%= l_rs("evafacnro")%></td>	
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="10%"><%= l_rs("falfecha")%></td>
        <td align="center" width="20%"><% if  l_rs("falpendiente") = 0 then %> Terminado <% else %> Pendiente <% End If %></td>
        <td align="center" width="10%"><%= l_rs("falporcen")%></td>		
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
