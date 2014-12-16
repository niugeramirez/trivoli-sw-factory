<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Berths.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: berth_con_excel.asp
'Descripción: Consulta de Berths
'Autor : Raul Chinestra
'Fecha: 17/03/2008

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_pornro

l_filtro = request("filtro")
l_orden  = request("orden")

l_pornro  = request("pornro")

if l_orden = "" then
  l_orden = " ORDER BY aredes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Berths - buques</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center">Berth</th>
    </tr>
	<tr>
	    <th>Descripci&oacute;n</th>    
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT berdes "
l_sql = l_sql & " FROM for_berth "
l_sql = l_sql & " WHERE for_berth.pornro = " & l_pornro
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td>No existen  Berths cargados.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="100%" nowrap><%= l_rs("berdes")%></td>
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
