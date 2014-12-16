<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: conexion_seg_00.asp
'Descripción: 
'Autor: Lisandro Moro
'Fecha: 15/03/2005
'Modificado:
on error goto 0

Dim rs
Dim sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cnnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Mantenimiento Conexiones - Supervisor - RHPro &reg;</title>
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
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
    </tr>
<%

Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT cnnro, cndesc "
sql = sql & "FROM conexion "
if l_filtro <> "" then
  sql = sql & "WHERE " & l_filtro & " "
end if
sql = sql & l_orden
rsOpen rs, cn, sql, 0 
do until rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= rs("cnnro")%>)">
        <td><%= rs("cnnro")%></td>
        <td><%= rs("cndesc")%></td>
    </tr>
<%
	rs.MoveNext
loop
rs.Close
set rs = Nothing
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
