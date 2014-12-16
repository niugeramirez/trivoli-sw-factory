<% Option Explicit
response.buffer = true
 %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: cupos_con_01.asp
'Descripción: Abm de  cupos
'Autor : Gustavo Manfrin
'Fecha: 29/12/2006

'on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_cont

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cupcod "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Cupos - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="center">Cupo</th>
		<th>C.Porte</th>
        <th>Producto</th>
		<th>Vendedor</th>
		<th>Corredor</th>
		<th>Fecha</th>
		<th>Empresa</th>
		<th>Operación</th>
    </tr>
<%
																	
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT cupcod, cupfec, tkt_movimiento.movcod, tkt_empresa.empcod, tkt_producto.prodes "
l_sql = l_sql & ", v.vencorcod vencod, v.vencordes vendes, c.vencorcod corcod, c.vencordes cordes "
l_sql = l_sql & ", tkt_cartaporte.carpornum "
l_sql = l_sql & " FROM tkt_cupo"
l_sql = l_sql & " INNER JOIN tkt_vencor v ON v.vencornro = tkt_cupo.vencornro "
l_sql = l_sql & " INNER JOIN tkt_vencor c ON c.vencornro = tkt_cupo.vencornro2 "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_producto.pronro = tkt_cupo.pronro "
l_sql = l_sql & " INNER JOIN tkt_empresa ON tkt_empresa.empnro = tkt_cupo.empnro "
l_sql = l_sql & " LEFT JOIN tkt_movimiento ON tkt_movimiento.movnro = tkt_cupo.movnro "
l_sql = l_sql & " LEFT JOIN tkt_cartaporte ON tkt_cartaporte.carpornro = tkt_movimiento.carpornro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="7">No existen Cupos</td>
</tr>
<%else
	l_cont = 0
	do until l_rs.eof
		l_cont = l_cont + 1
	%>
	    <tr>
	        <td width="15%" align="center"><%= l_rs("cupcod")%></td>
			<td width="15%" nowrap> <%= l_rs("carpornum")%></td>
	        <td width="15%" nowrap align="center" > <%= l_rs("prodes")%></td>
			<td width="30%" nowrap> <%=l_rs("vencod")%> - <%= l_rs("vendes")%></td>
			<td width="20%" nowrap> <%=l_rs("corcod")%> - <%= l_rs("cordes")%></td>
			<td width="10%" nowrap align="center"> <%= l_rs("cupfec")%></td>
			<td width="5%" nowrap align="center"> <%= l_rs("empcod")%></td>
			<td width="15%" nowrap> <%= l_rs("movcod")%></td>
	    </tr>
	<%
		if l_cont > 1000 then
			response.flush
			l_cont = 0
		end if
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
