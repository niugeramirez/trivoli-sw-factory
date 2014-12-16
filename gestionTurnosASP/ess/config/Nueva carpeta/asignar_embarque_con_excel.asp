<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Depositos.xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: asignar_cascara_con_excel.asp
'Descripci�n: Asignaci�n de Nros a los Camioneros para la C�scara
'Autor : Ra�l Chinestra	
'Fecha: 09/05/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY depdes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Dep�sitos - Ticket</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
        <th align="center" colspan="4">Dep�sitos</th>
    </tr>
    <tr>
        <th>C�digo</th>
        <th>Descripci&oacute;n</th>
        <th>Multiproducto</th>
        <th>Tipo</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT depcod,depdes,depmul,deptip"
l_sql = l_sql & " FROM tkt_deposito "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Dep�sitos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="20%" nowrap><%= l_rs("depcod")%></td>
	        <td width="40%" nowrap><%= l_rs("depdes")%></td>
	        <td width="10%" nowrap><%if l_rs("depmul") = -1 then%>Si<% Else %>No<% End If %></td>
	        <td width="30%" nowrap><%if UCase(l_rs("deptip")) = "C" then%>Celda/Silo<% Else %>Tanque<% End If %></td>
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
