<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Tipos de Buques.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: companies_con_excel.asp
'Descripción: Consulta de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY tipbuqdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Tipos de Buques - Buques</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
	    <th>Tipos de Buques</th>    
    </tr>
    <tr>
        <th>Código</th>
        <th nowrap>Descripción</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM buq_tipobuque "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="2" >No existen Tipos de Buques cargados.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="25%" nowrap><%= l_rs("tipbuqnro")%></td>
	        <td width="75%" nowrap><%= l_rs("tipbuqdes")%></td>			
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
