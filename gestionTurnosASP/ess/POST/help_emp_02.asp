<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: help_emp_02.asp
Descripción: Muestra los resultados de la busqueda del help_emp_01.asp
Autor : <Nombre del Creador>
Fecha: <Fecha de Creación>
Modificado:
			Fernando Favre - 22-07-03 - Se elimino la sql que calculaba el total de registros encontrados
			Fernando Favre - 31-07-03 - Se agrego empest en la consulta sql para poder diferenciar los empleados activos, inactivos o todos
-->
<% 
Dim l_rs
Dim l_sql
Dim l_empleg
Dim l_terape
Dim l_ternom

Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim total

l_filtro = request("filtro")
l_orden  = request("orden")

Dim l_estado
Dim l_selectdoc

l_estado  = request("estado")
l_selectdoc = request("selectdoc")

if l_orden = "" then
  l_orden = " Order by terape"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,legajo,apellido,nombre)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.empleg.value = legajo;
 document.datos.terape.value = apellido;
 document.datos.ternom.value = nombre;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="left">Empleado</th>
        <th align="left">Apellido</th>
        <th align="left">Nombre</th>
        <th align="center">Sigla</th>
        <th align="center">Nro. Documento</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'l_sql = "SELECT empleado.empleg, empleado.terape, empleado.ternom, empleado.empest, tipodocu.tidsigla, ter_doc.nrodoc, empleado.ternro "
'l_sql = l_sql & "FROM empleado, ter_doc, tipodocu "
'l_sql = l_sql & "WHERE empleado.ternro=ter_doc.ternro "
'l_sql = l_sql & "  and ter_doc.tidnro = (select min(tidnro) from ter_doc where ter_doc.ternro = empleado.ternro)"
'l_sql = l_sql & "  and ter_doc.tidnro=tipodocu.tidnro "

l_sql = "SELECT empleado.empleg, empleado.terape, empleado.ternom, tipodocu.tidsigla, ter_doc.nrodoc, empleado.ternro "
l_sql = l_sql & " FROM empleado"
l_sql = l_sql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro"
l_sql = l_sql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro"
l_sql = l_sql & " WHERE "
if l_selectdoc = 1 then
	l_sql = l_sql & " ter_doc.tidnro = (select min(tidnro) from ter_doc where ter_doc.ternro = empleado.ternro and ter_doc.tidnro < 4) "
else
	l_sql = l_sql & " ter_doc.tidnro=tipodocu.tidnro "
end if
if trim(l_estado) <> "" then
l_sql = l_sql & " and empleado.empest= " & l_estado
end if

if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if

l_sql = l_sql & l_orden

rsOpen l_rs, Cn, l_sql, 0

total = 0
do until l_rs.eof
total = total + 1
l_empleg = l_rs("empleg")
l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
%>
    <tr ondblclick="Javascript:parent.seleccionar();close();" onclick="Javascript:Seleccionar(this,<%= l_rs("ternro") %>,<%= l_empleg %>,'<%= l_terape %>','<%= l_ternom %>');">
        <td align="center"><%= l_empleg %></td>
        <td align="left"><%= l_terape %></td>
        <td align="left"><%= l_ternom %></td>
        <td align="center"><%= l_rs("tidsigla")%></td>
        <td align="center"><%= l_rs("nrodoc")%></td>
    </tr>
<%
	l_rs.MoveNext
loop
response.write "<script>parent.datos.total.value="&total&"</script>"

l_rs.close
l_rs = Nothing
cn.Close
cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="empleg" value="0">
<input type="Hidden" name="terape" value=" ">
<input type="Hidden" name="ternom" value=" ">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
