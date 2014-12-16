<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: ordenes_transportistas_con_01.asp
'Descripción: Consulta de Ordenes de trabajo
'Autor : Lisandro Moro
'Fecha: 22/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_ordnro

l_filtro = request("filtro")
l_orden  = request("orden")

l_ordnro = request.querystring("ordnro")

if l_orden = "" then
  l_orden = " ORDER BY tracod "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>
<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Ordenes de trabajo - Ticket</title>
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
	parent.document.all.camiones.src = 'ordenes_camioneros_con_01.asp?ordnro=<%= l_ordnro %>&tranro=' + cabnro;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Código</th>
        <th>Descripción</th>
        <th>Razón Social</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Solo mostraré las órdenes vigentes
l_sql = "SELECT tkt_transportista.tranro, tracod, trades, trarazsoc "
l_sql = l_sql & " FROM tkt_transportista "
l_sql = l_sql & " LEFT JOIN tkt_ord_tra ON tkt_transportista.tranro = tkt_ord_tra.tranro"
l_sql = l_sql & " WHERE ordnro = " & l_ordnro
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Transportistas asociados a la Orden de trabajo</td>
</tr>
<%else
	do until l_rs.eof
	%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("tranro")%>)">
        <td width="10%" nowrap align="center"><%= l_rs("tracod")%></td>
        <td width="20%" nowrap><%= l_rs("trades")%></td>
        <td width="25%" nowrap><%= l_rs("trarazsoc")%></td>
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
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
