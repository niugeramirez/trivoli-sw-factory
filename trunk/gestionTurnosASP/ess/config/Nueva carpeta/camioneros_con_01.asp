<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: camioneros_con_01.asp
'Descripción: Abm de camioneros
'Autor : Lisandro Moro
'Fecha: 15/02/2005

on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_tipo

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY camdes "
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
<title><%= Session("Titulo")%>Camioneros - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,des){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.descripcion.value = des;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
		<th>Código</th>
        <th>Apellido y Nombre</th>
		<th>Chasis</th>
		<th>Acoplado</th>
		<th>Habilitado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  camnro, camcod, camdes, camhab, camsis, camcha, camaco "'trades, 
l_sql = l_sql & " FROM tkt_camionero "
'l_sql = l_sql & " LEFT JOIN tkt_transportista ON tkt_transportista.tranro = tkt_camionero.tranro "
l_sql = l_sql & " WHERE camact = -1 "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Camioneros</td>
</tr>
<%else
	do until l_rs.eof	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('camioneros_con_02.asp?tipo=<% If l_rs("camsis") = -1 then %>C<% Else %>M<% End If %>&cabnro=' + datos.cabnro.value,'',500,400)" onclick="Javascript:Seleccionar(this,<%= l_rs("camnro")%>,'<%= l_rs("camdes")%>')">
	        <td width="10%" nowrap align="center"><%= l_rs("camcod")%></td>
			<td width="60%" nowrap><%= l_rs("camdes")%></td>
			<td width="15%" nowrap align="center"><%= l_rs("camcha") %></td>
			<td width="15%" nowrap align="center"><%= l_rs("camaco") %></td>
			<td width="10%" nowrap align="center"><%if l_rs("camhab") = -1 then %>Si<% Else %>No<% End If %></td>
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
<input type="hidden" name="descripcion" value="">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
