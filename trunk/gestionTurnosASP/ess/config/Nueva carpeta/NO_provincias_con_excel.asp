<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Provincias.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: provincias_con_excel.asp
'Descripción: Consulta de Provincias
'Autor : Alvaro Bayon
'Fecha: 08/02/2005
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
  l_orden = " ORDER BY pronro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Provincias - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="2">Provincias</th>
    </tr>
    <tr>
        <th>Código</th>		
        <th>Descripci&oacute;n</th>		
        <th>Oblea</th>				
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT pronro, procod, prodes, proobl "
l_sql = l_sql & " FROM tkt_provincia "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="1">No existen Provincias</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td  align="center" width="20%" nowrap><%= l_rs("procod")%></td>
	        <td width="70%" nowrap><%= l_rs("prodes")%></td>
	        <td align="center" width="10%" nowrap><% if l_rs("proobl") = -1 then %>Si <% Else  %>No<% End If %></td>			
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
