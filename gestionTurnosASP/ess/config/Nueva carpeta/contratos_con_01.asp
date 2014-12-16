<% Option Explicit
response.buffer = true
 %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: contratos_con_01.asp
'Descripción: Abm de  contratos
'Autor : Lisandro Moro
'Fecha: 11/02/2005

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
  l_orden = " ORDER BY concod "
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
<title><%= Session("Titulo")%>Contratos - Ticket</title>
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
        <th align="center">Código</th>
        <th>Producto</th>
		<th>Vendedor</th>
		<th>Kgs</th>
		<th>Fecha</th>
    </tr>
<%
																	
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT connro, concod, confec, prodes, vencordes, conkil "
l_sql = l_sql & " FROM tkt_contrato "
l_sql = l_sql & " INNER JOIN tkt_vencor ON tkt_vencor.vencornro = tkt_contrato.vennro "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_producto.pronro = tkt_contrato.pronro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Contratos</td>
</tr>
<%else
	l_cont = 0
	do until l_rs.eof
		l_cont = l_cont + 1
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('contratos_con_02.asp?cabnro=' + datos.cabnro.value,'',520,500)" onclick="Javascript:Seleccionar(this,<%= l_rs("connro")%>)">
	        <td width="20%" align="center"><%= l_rs("concod")%></td>
	        <td width="20%" nowrap><%= l_rs("prodes")%></td>
			<td width="30%" nowrap><%= l_rs("vencordes")%></td>
			<td width="15%" nowrap align="right"><%= l_rs("conkil")%></td>
			<td width="20%" nowrap align="center"><%= l_rs("confec")%></td>
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
