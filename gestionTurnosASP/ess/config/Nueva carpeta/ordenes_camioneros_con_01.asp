<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: ordenes_camioneros_con_01.asp
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

Dim l_tranro
Dim l_ordnro

l_filtro = request("filtro")
l_orden  = request("orden")

l_tranro = request.querystring("tranro")
l_ordnro = request.querystring("ordnro")

if l_orden = "" then
  l_orden = " ORDER BY camcod "
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
<title><%= Session("Titulo")%>Camioneros en Ordenes de trabajo - Ticket</title>
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
        <th>Descripción</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")


'Solo mostraré las órdenes vigentes
l_sql = "SELECT tkt_camionero.camnro, camcod, camdes "
l_sql = l_sql & " FROM tkt_camionero "
l_sql = l_sql & " LEFT JOIN tkt_ord_cam ON tkt_camionero.camnro = tkt_ord_cam.camnro "
l_sql = l_sql & " LEFT JOIN tkt_cam_tra ON tkt_camionero.camnro = tkt_cam_tra.camnro "
l_sql = l_sql & " WHERE tkt_cam_tra.tranro = " & l_tranro
l_sql = l_sql & " AND tkt_ord_cam.ordnro = " & l_ordnro
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	if l_tranro = 0 and l_ordnro = 0 then%>
	<tr>
		 <td colspan="5">Seleccione un Transportista</td>
	</tr>
	<% Else %>
	<tr>
		 <td colspan="5">No existen camioneros asociados al Transportista en esta Orden de trabajo</td>
	</tr>
	<% End If %>
<%else
	do until l_rs.eof
	%>
    <tr ondblclick="Javascript:parent.abrirVentana('camioneros_con_02.asp?tipo=C&cabnro=' + datos.cabnro.value,'',600,400)" onclick="Javascript:Seleccionar(this,<%= l_rs("camnro")%>)">
        <td width="10%" nowrap align="center"><%= l_rs("camcod")%></td>
        <td width="20%" nowrap><%= l_rs("camdes")%></td>
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
