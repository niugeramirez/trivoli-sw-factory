<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Entregadores - Recibidores.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: entrec_con_excel.asp
'Descripción: Abm de Entregadores y recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY entdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Entregadores/Recibidores - ticket - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="3">Entregadores/Recibidores</th>
    </tr>
    <tr>
        <th align="center">C&oacute;digo</th>
        <th>Descripci&oacute;n</th>
        <th>Tipo</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT entnro,entcod,entdes,entact,entrol"
l_sql = l_sql & " FROM tkt_entrec"
l_sql = l_sql & " WHERE entact = -1"
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Entregadores/Recibidores</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="10%" align="center"><%= l_rs("entcod")%></td>
	        <td width="60%" nowrap><%= l_rs("entdes")%></td>
	        <td width="30%" align="center" nowrap><% if l_rs("entrol")="E" then%>Entregador<%else if l_rs("entrol")="A" then%>Ambos<%else%>Recibidor<%end if end if%></td>
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
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
