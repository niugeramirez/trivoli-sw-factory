<%  Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% Response.AddHeader "Content-Disposition", "attachment;filename=Politicas de Cuentas.xls" %>
<%

'Archivo: pol_cuenta_seg_excel.asp
'Descripción: ABM de Políticas de cuenta
'Autor: Alvaro Bayon
'Fecha: 21/02/2005

 Dim l_rs
 Dim l_sql
 
 Dim l_filtro
 Dim l_orden
 
 l_filtro = request("filtro")
 l_orden  = request("orden")
 
 if l_orden = "" then
	l_orden = "ORDER BY pol_desc"
 end if
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Pol&iacute;ticas de Cuentas - Ticket</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th colspan="1">Pol&iacute;ticas de Cuentas</th>
    </tr>
    <tr>
        <th >Descripci&oacute;n</th>
    </tr>
<%
 l_sql = "SELECT pol_nro, pol_desc "
 l_sql = l_sql & "FROM pol_cuenta "
 if l_filtro <> "" then
 	l_sql = l_sql & "WHERE " & l_filtro & " "
 end if
 l_sql = l_sql & l_orden
 rsOpen l_rs, cn, l_sql, 0 
 
 do until l_rs.eof
	%>
    <tr>
        <td><%= l_rs("pol_desc")%></td>
    </tr>
	<%
	l_rs.MoveNext
 loop
 
 l_rs.Close
 cn.Close
 set l_rs = Nothing 
 set cn = Nothing 
%>
</table>
</body>
</html>
