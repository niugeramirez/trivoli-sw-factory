<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: ordenes_con_01.asp
'Descripción: Consulta de Ordenes de trabajo
'Autor : Alvaro Bayon
'Fecha: 09/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY ordcod "
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
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Código</th>
        <th>Producto</th>
        <th>Origen</th>
        <th>Destino</th>
        <th>Habilitado</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Solo mostraré las órdenes vigentes
l_sql = "SELECT ordnro, ordcod, prodes, tkt_lugar.lugdes, lugard.lugdes as deslugdes, ordhab "
l_sql = l_sql & " FROM tkt_ordentrabajo "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_ordentrabajo.pronro = tkt_producto.pronro"
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_ordentrabajo.orilugnro = tkt_lugar.lugnro"
l_sql = l_sql & " INNER JOIN tkt_lugar lugard  ON tkt_ordentrabajo.deslugnro = lugard.lugnro"
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Ordenes de trabajo</td>
</tr>
<%else
	do until l_rs.eof
	%>
    <tr ondblclick="Javascript:parent.abrirVentana('ordenes_con_02.asp?cabnro=' + datos.cabnro.value,'',680,500)" onclick="Javascript:Seleccionar(this,<%= l_rs("ordnro")%>)">
        <td width="10%" nowrap align="center"><%= l_rs("ordcod")%></td>
        <td width="20%" nowrap><%= l_rs("prodes")%></td>
        <td width="25%" nowrap><%= l_rs("lugdes")%></td>
        <td width="25%" nowrap><%= l_rs("deslugdes")%></td>
        <td width="10%" align="center" nowrap><% if l_rs("ordhab")=-1 then%>Si<%else%>No<%end if%></td>
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
