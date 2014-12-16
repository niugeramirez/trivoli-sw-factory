<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=vendedor.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: vendedores_con_excel.asp
'Descripción: Abm de Vendedores
'Autor : Gustavo Manfrin
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
  l_orden = " ORDER BY vencornro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="4">Vendedores</th>
    </tr>
    <tr>
        <th align="center">C&oacute;digo</th>
        <th>Nombre clave</th>		
        <th>Razón social</th>		
        <th>Hab.</th>				
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT vencornro, vencorcod, vencordes, vencorrazsoc, venhab "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND vencortip = 'V' "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Vendedores</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="15%" align="center"><%= l_rs("vencorcod")%></td>
	        <td width="25%" nowrap><%= l_rs("vencordes")%></td>
	        <td width="50%" nowrap><%= l_rs("vencorrazsoc")%></td>			
            <td width="10%" align="center"><% if l_rs("venhab") = -1 then%> Si <% Else %> No <% End If%></td>			
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
