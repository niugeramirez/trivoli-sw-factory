<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->

<%
' Modificado: 14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion	
'		    	18-08-2006 - LA. - Sacar la vista v_empleado
%>
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

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY terape "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Formulario de Carga - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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
        <th align="center"><%if ccodelco=-1 then%>N&uacute;mero<%else%>Empleado<%end if%></th>
        <th align="left">Apellido</th>
        <th align="left">Nombre</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado.empleg, empleado.terape, empleado.ternom, empleado.ternro FROM empleado INNER JOIN evacab ON empleado.ternro=evacab.empleado "
if l_filtro <> "" then
  l_sql = l_sql & " Where " & l_filtro & " "
end if
l_sql = l_sql & l_orden

rsOpen l_rs, Cn, l_sql, 0
do until l_rs.eof
l_empleg = l_rs("empleg")
l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
%>
    <tr ondblclick="Javascript:parent.seleccionar();close();" onclick="Javascript:Seleccionar(this,<%= l_rs("ternro") %>,<%= l_empleg %>,'<%= l_terape %>','<%= l_ternom %>');">
        <td align="center"><%= l_empleg %></td>
        <td align="left"><%= l_terape %></td>
        <td align="left"><%= l_ternom %></td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close

l_sql = "select COUNT(ternro) as cant FROM empleado INNER JOIN evacab ON empleado.ternro=evacab.empleado "
if l_filtro <> "" then
  l_sql = l_sql & " Where " & l_filtro & " "
end if
rsOpen l_rs, cn, l_sql, 0
response.write "<script>parent.datos.total.value="&l_rs("cant")&"</script>"

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
<input type="Hidden" name="orden" value="<%=l_orden %>">
<input type="Hidden" name="filtro" value="<%=l_filtro %>">
</form>
</body>
</html>
