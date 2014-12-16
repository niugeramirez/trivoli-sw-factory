<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: camioneros_transportistas_con_01.asp
'Descripci�n: Consulta de transportistas asociados al camionero
'Autor : Lisandro Moro
'Fecha: 22/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_camnro
Dim l_ordnro

l_filtro = request("filtro")
l_orden  = request("orden")
l_camnro = request.querystring("camnro")

if l_orden = "" then
  l_orden = " ORDER BY tracod "
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
        <th>C�digo</th>
        <th>Descripci�n</th>
        <th>Raz�n Social</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Solo mostrar� las �rdenes vigentes
l_sql = "SELECT tkt_transportista.tranro, tracod, trades, trarazsoc "
l_sql = l_sql & " FROM tkt_transportista "
l_sql = l_sql & " LEFT JOIN tkt_cam_tra ON tkt_transportista.tranro = tkt_cam_tra.tranro"
l_sql = l_sql & " WHERE camnro = " & l_camnro
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Transportistas asociados al Camionero</td>
</tr>
<%else
	do until l_rs.eof
	%>
    <tr>
        <td width="10%" nowrap><%= l_rs("tracod")%></td>
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
