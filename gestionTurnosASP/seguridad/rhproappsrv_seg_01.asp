<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: rhproappserv_seg_00.asp
'Descripción: Se encarga de ejecutar los procesos del sistema
'Autor : Lisandro Moro
'Fecha : 11/03/2005
'Modificado:

Dim rs
Dim sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request.QueryString("orden")

if l_orden = "" then
  l_orden = " ORDER BY cabpolnivel ASC"  'orden por defecto nivel asc
end if

dim l_cabpolnro
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>RHPro AppServer</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};

	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Id</th>
        <th>Nombre del Servicio</th>
    </tr>
<%

Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT id, servdesabr "
sql = sql & "FROM rhproappsrv ORDER BY servdesabr"
rsOpen rs, cn, sql, 0 

do until rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= rs("id")%>)">
        <td><%= rs("id")%></td>
        <td><%= rs("servdesabr")%></td>
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
</form>
</body>
</html>
