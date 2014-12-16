<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=impresoras.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: impresoras_con_excel.asp
'Descripción: ABM de Impresoras
'Autor : Lisandro Moro
'Fecha: 26/09/2005
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
  l_orden = " ORDER BY tkt_impresora.impnom "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Impresoras - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Nombre Impresora</th>
        <th>Nombre Compartido</th>
        <th>Matricial</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT impnro,  impnom, impnomcom, impmat "
l_sql = l_sql & " FROM tkt_impresora "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen impresoras</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="20%" nowrap><%= l_rs("impnom")%></td>
	        <td width="20%" nowrap><%= l_rs("impnomcom")%></td>
	        <td width="40%" nowrap><%if l_rs("impmat") = 0 then%>No<% Else %>Si<% End If %></td>
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
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>