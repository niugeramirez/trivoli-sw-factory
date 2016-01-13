<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim rs
Dim sql
Dim l_filtro
Dim l_filtro2
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request.Querystring("filtro")
l_orden  = request.QueryString("orden")

l_filtro = request("filtro") 


if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	


if l_orden = "" then
  l_orden = " ORDER BY mulnro ASC"  'orden por defecto número desc
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuración multiemprea - Ticket</title>
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
        <th>Número</th>
        <th>Nombre</th>
        <th>Múltiple</th>
    </tr>
<%

Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT mulnro,mulnom,multiple "
sql = sql & "FROM multiempresa "
if l_filtro <> "" then
  sql = sql & "WHERE " & l_filtro & " "
end if
sql = sql & l_orden
rsOpen rs, cn, sql, 0 
do until rs.eof
%>
    <tr ondblclick="Javascript:parent.abrirVentana('multiempresa_gti_02.asp?Tipo=M&mulnro=' + datos.cabnro.value,'',400,100)" onclick="Javascript:Seleccionar(this,<%= rs("mulnro")%>)">
        <td><%= rs("mulnro")%></td>
        <td><%= rs("mulnom")%></td>
        <td><% 
		if rs("multiple") = 0 then
			Response.write("No")
		else
			Response.write("Si")
		end if
		%></td>
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
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">
</form>
</body>
</html>
