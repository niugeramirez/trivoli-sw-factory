<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Tipos de Documentos.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: tipos_documentos_con_excel.asp
'Descripción: Abm de Tipos de Documentos
'Autor : Raul Chinestra
'Fecha: 08/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY tipdocdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipos de Documentos - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="2">Tipos de Documentos</th>
    </tr>
    <tr>
        <th align="center">Sigla</th>
        <th>Descripci&oacute;n</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tipdocnro, tipdocsig, tipdocdes "
l_sql = l_sql & " FROM tkt_tipodocumento "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Tipos de Documentos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="20%" align="center"><%= l_rs("tipdocsig")%></td>
	        <td width="80%" nowrap><%= l_rs("tipdocdes")%></td>
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
