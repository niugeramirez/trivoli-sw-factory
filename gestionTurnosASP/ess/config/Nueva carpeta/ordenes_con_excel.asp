<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Ordenes de Trabajo.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: ordenes_con_excel.asp
'Descripción: Consulta de Ordenes de trabajo
'Autor : Alvaro Bayon
'Fecha: 09/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY ordcod "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Ordenes de Trabajo- Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="5">Ordenes de Trabajo</th>
    </tr>
    <tr>
        <th>Código</th>
        <th>Producto</th>
        <th>Origen</th>
        <th>Destino</th>
        <th>Habilitado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ordnro, ordcod, prodes, tkt_lugar.lugdes, lugard.lugdes as deslugdes, ordhab "
l_sql = l_sql & " FROM tkt_ordentrabajo "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_ordentrabajo.pronro = tkt_producto.pronro"
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_ordentrabajo.orilugnro = tkt_lugar.lugnro"
l_sql = l_sql & " INNER JOIN tkt_lugar lugard  ON tkt_ordentrabajo.deslugnro = lugard.lugnro"
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Ordenes de trabajo</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
        <td width="10%" align="center" nowrap><%= l_rs("ordcod")%></td>
        <td width="20%" nowrap><%= l_rs("prodes")%></td>
        <td width="25%" nowrap><%= l_rs("lugdes")%></td>
        <td width="25%" nowrap><%= l_rs("deslugdes")%></td>
        <td width="20%" align="center" nowrap><% if l_rs("ordhab")=-1 then%>Si<%else%>No<%end if%></td>
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
