<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_cystipnro
Dim l_cystipact 
Dim l_cystipsis
Dim l_cystipmsg
Dim l_cystipmail

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cystipnro"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipo de Autorizaci&oacute;n - Ticket</title>
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
        <th>Nombre</th>
        <th>Activo</th>
        <th>Mensajes</th>
        <th>Mail</th>
        <th>Sistema</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT cystipnro, cystipnombre, cystipact, cystipmsg, cystipmail, cystipsis "
l_sql = l_sql & "FROM cystipo "
if l_filtro <> "" then
  l_sql = l_sql & "WHERE " & l_filtro & " "
end if
l_sql = l_sql & l_orden
l_rs.Maxrecords = 50
rsOpen l_rs, cn, l_sql, 0 

l_cystipnro = l_rs("cystipnro")

do until l_rs.eof
	' poner campo binary con formato si/no en browse
	if l_rs("cystipact") then
		l_cystipact = "si"
	else
		l_cystipact = "no"
	end if
	if l_rs("cystipsis") then
		l_cystipsis = "si"
	else
		l_cystipsis = "no"
	end if
	if l_rs("cystipmsg") then
		l_cystipmsg = "si"
	else
		l_cystipmsg = "no"
	end if
	if l_rs("cystipmail") then
		l_cystipmail = "si"
	else
		l_cystipmail = "no"
	end if



%>
    <tr ondblclick="Javascript:parent.abrirVentana('autorizacion_gti_02.asp?Tipo=M&cystipnro=' + datos.cabnro.value,'',700,300)" onclick="Javascript:Seleccionar(this,<%= l_rs("cystipnro")%>)">
        <td><%= l_rs("cystipnro")%></td>
        <td><%= l_rs("cystipnombre")%></td>
        <td><%= l_cystipact%></td>
        <td><%= l_cystipmsg%></td>
        <td><%= l_cystipmail%></td>
        <td><%= l_cystipsis%></td>
    </tr>
<%
	l_rs.MoveNext
loop

l_rs.Close
cn.Close
set l_rs = Nothing 
set cn = Nothing 
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="<%= l_cystipnro %>" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
