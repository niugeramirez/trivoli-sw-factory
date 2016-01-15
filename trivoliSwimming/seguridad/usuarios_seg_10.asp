<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 
on error goto 0
Dim l_rs
Dim l_sql
Dim l_tenro


Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_iduser

l_iduser = request("iduser")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Estructuras del Usuario - Ticket</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "select tipoestructura.tenro, tedabr, estrdabr "
l_sql = l_sql & "from usupuedever, estructura,tipoestructura "
l_sql = l_sql & "where iduser = '" & l_iduser & "' "
l_sql = l_sql & "and usupuedever.tenro = tipoestructura.tenro "
l_sql = l_sql & "and estructura.estrnro = usupuedever.estrnro "
l_sql = l_sql & "order by tedabr"

l_rs.Maxrecords = 100
rsOpen l_rs, cn, l_sql, 0
l_tenro = 0
if l_rs.eof then
	%>
	<tr nowrap>
		<td><b>No se encontraron datos</b></td>
	</tr>
	<%	
else
	do until l_rs.eof
		if l_tenro <> l_rs("tenro") then %>
	    <tr nowrap>
	        <th align="center"><%=  l_rs("tedabr")%></th>
	    </tr>
	<%	
		l_tenro = l_rs("tenro")
		end if
	%>
	    <tr>
	        <td align="center"><%= l_rs("estrdabr") %></td>
	    </tr>
	<%
		l_rs.MoveNext
	loop
end if
l_rs.Close

Set l_rs = Nothing
cn.Close
Set cn = Nothing
%>
</table>
</body>
</html>
