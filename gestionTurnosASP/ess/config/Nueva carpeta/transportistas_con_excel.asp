<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Transportistas.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: transportistas_con_excel.asp
'Descripción: Abm de Transportistas
'Autor : Gustavo Manfrin
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
  l_orden = " ORDER BY tranro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Transportistas - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="8">Transportistas</th>
    </tr>
    <tr>
        <th align="center">C&oacute;digo</th>
        <th>Nombre clave</th>		
        <th>Razón social</th>		
        <th>Direccion</th>				
		<th>Provincia</th>		
		<th>Caja</th>		
		<th>Nro.Caja</th>		
		<th>Situacion impositiva</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tranro, tracod, trades, trarazsoc, tracaj, tranrocaj, tradir, sitimpdes, prodes "
l_sql = l_sql & " FROM tkt_transportista "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_transportista.pronro  "
l_sql = l_sql & " INNER JOIN tkt_sitimp ON tkt_sitimp.sitimpnro = tkt_transportista.sitimpnro  "
l_sql = l_sql & " WHERE traact = -1 "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="8">No existen Transportistas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="5%" align="center"><%= l_rs("tracod")%></td>
	        <td width="10%" nowrap><%= l_rs("trades")%></td>
	        <td width="30%" nowrap><%= l_rs("trarazsoc")%></td>			
            <td width="10%" nowrap><%= l_rs("tradir")%></td>			
            <td width="10%" nowrap><%= l_rs("prodes")%></td>						
            <td width="7%" nowrap><%= l_rs("tracaj")%></td>			
            <td width="7%" nowrap><%= l_rs("tranrocaj")%></td>			
            <td width="10%" nowrap><%= l_rs("sitimpdes")%></td>			
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
