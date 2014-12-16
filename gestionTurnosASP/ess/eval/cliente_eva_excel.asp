<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% Response.AddHeader "Content-Disposition", "attachment;filename=clientes.xls" %>
<%
'================================================================================
'Archivo		: cliente_eva_excel.asp
'Descripción	: Excel de Clientes
'Autor			: CCRossi
'Fecha			: 13-12-2004 
'Modificado		: 
'================================================================================

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY evaclinro "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Clientes - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th colspan=3>Clientes</th>
    </tr>
    <tr>
        <th>C&oacute;digo</th>
        <th>C&oacute;d.Ext.</th>
        <th>Raz&oacute;n Social</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaclinro, evaclinom ,evaclicodext "
l_sql = l_sql & " FROM evacliente "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No hay Clientes.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="10%" align="right"><%= l_rs("evaclinro")%></td>
	        <td width="20%" nowrap align="left"><%= l_rs("evaclicodext")%></td>
	        <td width="45%" nowrap><%= l_rs("evaclinom")%></td>
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
</body>
</html>
