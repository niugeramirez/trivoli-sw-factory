<% Option Explicit %>
<%
Response.AddHeader "Content-Disposition", "attachment;filename=confempr.xls" 
Response.ContentType = "application/vnd.ms-excel" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : conf_x_empr_excel.asp
Descripcion    : Modulo que se encarga de listar la configuracion por empresa.
Creador        : Scarpa D.
Fecha Creacion : 01/10/2003
-----------------------------------------------------------------------------
-->

<% 
Dim l_rs
Dim l_sql

Dim l_repnro
Dim l_confnrocol
Dim l_confetiq
Dim l_conftipo
dim	l_confval

Dim l_filtro
Dim l_orden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY confper.confnro ASC"  'orden por número asc
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Empresas - Ticket</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
        <th>Activo</th>
        <th>Valor</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM  confper "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if

l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	<tr>
		<td width="10%"><%=l_rs("confnro")%></td>
		<td width="40%"><%=l_rs("confdesc")%> </td>
		<td width="25%" align="center"><% if CInt(l_rs("confactivo")) = -1 then response.write "Activo" else response.write "Inactivo" end if%> </td>
		<td width="25%" align="center"><%=l_rs("confint")%> </td>
	</tr>
	<%l_rs.MoveNext
	loop
end if 
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>

</body>
</html>
