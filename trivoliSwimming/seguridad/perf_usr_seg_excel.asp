<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% Response.AddHeader "Content-Disposition", "attachment;filename=Perfiles de Usuarios.xls" %>
<%

'Archivo: perf_usr_seg_excel.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY perfnom "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Perfiles de Usuarios - Ticket</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th colspan="3">Perfiles de Usuarios</th>
    </tr>
    <tr>
        <th align="left">Descripci&oacute;n</th>
        <th>Pol&iacute;tica de Cuenta</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT perfnro, perfnom, perftipo, pol_desc "
l_sql = l_sql & "FROM perf_usr LEFT JOIN pol_cuenta ON perf_usr.pol_nro = pol_cuenta.pol_nro "
if l_filtro <> "" then
  l_sql = l_sql & "WHERE " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
%>
    <tr>
        <td align="left"><%= l_rs("perfnom")%></td>
        <td><%= l_rs("pol_desc")%></td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
</body>
</html>
