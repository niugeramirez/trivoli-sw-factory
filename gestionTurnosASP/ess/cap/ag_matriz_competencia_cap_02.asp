<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: ag_matriz_competencias_cap_01.asp
Descripcion: Gap matriz por Competencias
Autor: Lisandro Moro
Fecha: 29/03/2004
Modificado:
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql

Dim l_ternro
Dim l_orinro
Dim l_estnro

Dim l_saltear

l_ternro = l_ess_ternro

l_estnro = request.querystring("estado")

l_orinro = request.querystring("origen")
if l_orinro = "" then
	l_orinro= "0"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_estiloTabla%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gap Registrados por Competencias - Capacitaci�n - RHPro &reg;</title>
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
	    <th align="center">C�digo</th>
	    <th align="center">Fecha</th>
    	<th align="left">Competencia</th>
        <th align="center">Estado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evafacnro, falorigen, evafacdesabr, falidnro, falfecha, falpendiente, falporcen"
l_sql = l_sql & " FROM cap_falencia "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = cap_falencia.modnro "
l_sql = l_sql & " WHERE ternro = " & l_ternro 
l_sql = l_sql & " AND falorigen = 7 "
'l_sql = l_sql & " AND (falorigen = " & l_orinro & " OR 0 = " & l_orinro & ")"

if CStr(l_estnro) <> "2" then 
	l_sql = l_sql & " AND falpendiente = " & l_estnro
end if 

l_sql = l_sql & " ORDER BY falorigen"

rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then%>
	<tr>
        <td colspan=5>No hay Gap registrados por Competencias</td>
    </tr>
<%else
do until l_rs.eof
%>
	<tr onclick="Javascript:Seleccionar(this,<%= l_rs("evafacnro")%>)">
        <td align="center" width="10%"><%= l_rs("evafacnro")%></td>
        <td align="center" width="10%"><%= l_rs("falfecha")%></td>
        <td align="left" width="50%"><%= l_rs("evafacdesabr")%></td>
        <td align="center" width="20%"><% if  l_rs("falpendiente") = 0 then %> Terminado <% else %> Pendiente <% End If %></td>
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
<input type="hidden" name="cabnro" value="0">
</form>
</body>
</html>
