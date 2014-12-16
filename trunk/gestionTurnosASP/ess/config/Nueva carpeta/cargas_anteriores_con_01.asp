
 <% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: cargas_anteriores_con_01.asp
'Descripción: Abm de cargas anteriores
'Autor : Gustavo Manfrin
'Fecha: 07/08/2006
'Modificado: 

on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY lugdes,tkt_producto.prodes "
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
<title><%= Session("Titulo")%>Cargas Anteriores - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,lugnro,pronro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.lugnro.value = lugnro;
	document.datos.pronro.value = pronro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="center">Lugar Destino</th>
        <th>Producto Cargado</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT tkt_cargasconf.carconnro,tkt_cargasconf.lugdesnro, tkt_cargasconf.pronro, "
l_sql = l_sql & " tkt_lugar.lugcod, tkt_producto.prodes "
l_sql = l_sql & " FROM tkt_cargasconf "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_cargasconf.lugdesnro= tkt_lugar.lugnro "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_cargasconf.pronro= tkt_producto.pronro "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="7">No existen datos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('cargas_anteriores_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value + '&pronro=' + datos.pronro.value + '&pronro1=' + datos.pronro1.value,'',550,200)" onclick="Javascript:Seleccionar(this,<%= l_rs("carconnro")%>,<%= l_rs("lugdesnro")%>,<%= l_rs("pronro")%>)">
	        <td width="20%" align="center"><%= l_rs("lugcod")%></td>
	        <td width="80%" nowrap><%= l_rs("prodes")%></td>
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
<input type="hidden" name="lugnro" value="0">
<input type="hidden" name="pronro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
