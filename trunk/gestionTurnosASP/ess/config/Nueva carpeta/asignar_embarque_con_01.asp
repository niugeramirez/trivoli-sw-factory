<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: asignar_embarque_con_01.asp
'Descripción: Asignación de Nros a los Camioneros para el embarque
'Autor : Gustavo Manfrin	
'Fecha: 19/09/2006

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY tarcod "
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
<title><%= Session("Titulo")%>Asignar Nros para Embarque - Ticket</title>
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
        <th>Tarjeta Nro.</th>
        <th>Orden de Trabajo</th>		
        <th>Camionero</th>
        <th>Transportista</th>		
        <th>Embarque</th>				
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tkt_asiemb.asiembnro, tkt_asiemb.tarcod, tkt_ordentrabajo.ordcod, "
l_sql = l_sql & " tkt_camionero.camdes, tkt_transportista.trades, tkt_embarque.embcod"
l_sql = l_sql & " FROM tkt_asiemb "
l_sql = l_sql & " INNER JOIN tkt_camionero     ON tkt_camionero.camnro = tkt_asiemb.camnro " 
l_sql = l_sql & " INNER JOIN tkt_transportista ON tkt_transportista.tranro = tkt_asiemb.tranro " 
l_sql = l_sql & " INNER JOIN tkt_embarque      ON tkt_embarque.embnro = tkt_asiemb.embnro " 
l_sql = l_sql & " INNER JOIN tkt_ordentrabajo  ON tkt_ordentrabajo.ordnro = tkt_embarque.ordnro " 

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen datos.</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('asignar_embarque_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,450)" onclick="Javascript:Seleccionar(this,<%= l_rs("asiembnro")%>)">
	        <td width="20%" nowrap align="center"><%= l_rs("tarcod")%></td>
	        <td width="20%" nowrap align="center"><%= l_rs("ordcod")%></td>								
	        <td width="20%" nowrap align="center"><%= l_rs("camdes")%></td>
	        <td width="20%" nowrap align="center"><%= l_rs("trades")%></td>
	        <td width="20%" nowrap align="center"><%= l_rs("embcod")%></td>			
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
