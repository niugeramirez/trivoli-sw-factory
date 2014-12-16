<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--
Archivo: help_emp_02.asp
Descripción: Muestra los resultados de la busqueda del help_emp_01.asp
Autor : <Nombre del Creador>
Fecha: <Fecha de Creación>
Modificado:
			Fernando Favre - 22-07-03 - Se elimino la sql que calculaba el total de registros encontrados
			Fernando Favre - 31-07-03 - Se agrego empest en la consulta sql para poder diferenciar los empleados activos, inactivos o todos
Modificado: 19-11-04 CCRossi control de caracteres raros
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

if l_orden = "" then
  l_orden = " Order by terape"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
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
        <th align="left"><%if ccodelco=-1 then%>Supervisado<%else%>Empleado<%end if%></th>
        <th align="left">Apellido</th>
        <th align="left">Nombre</th>
        <th align="center">Sigla</th>
        <th align="center">Nro. Documento</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado.empleg, empleado.terape, empleado.ternom, empleado.empest, tipodocu.tidsigla, ter_doc.nrodoc, empleado.ternro FROM empleado, ter_doc, tipodocu WHERE empleado.ternro=ter_doc.ternro and ter_doc.tidnro = (select min(tidnro) from ter_doc where ter_doc.ternro = empleado.ternro) and ter_doc.tidnro=tipodocu.tidnro "
if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, Cn, l_sql, 0
'Response.Write l_sql
total = 0
if l_rs.EOF then%>
	<tr>
        <td colspan=5>No hay resultados para el Filtro.</td>
    </tr>
<%else
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
end if
response.write "<script>parent.datos.total.value="&total&"</script>"

l_rs.close
set l_rs = Nothing
cn.Close
set cn = Nothing
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
