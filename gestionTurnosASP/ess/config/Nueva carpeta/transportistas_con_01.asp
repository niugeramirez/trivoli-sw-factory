<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: transportistas_con_01.asp
'Descripción: Abm de Transportistas
'Autor : Gustavo Manfrin
'Fecha: 09/02/2005
'Modificado: 

'on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY trades "
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
<title><%= Session("Titulo")%>Transportistas - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro, des){
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
        <th align="center">C&oacute;digo</th>
        <th>Nombre clave</th>		
        <th>Razón social</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT tranro, tracod, trades, trarazsoc "
l_sql = l_sql & " FROM tkt_transportista "
l_sql = l_sql & " WHERE traact = -1 "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
l_rs.maxrecords = 4000
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Transportistas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('transportistas_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',680,300)" onclick="Javascript:Seleccionar(this,<%= l_rs("tranro")%>,'<%= l_rs("trades")%>')">
	        <td width="20%" align="center"><%= l_rs("tracod")%></td>
	        <td width="30%" nowrap><%= l_rs("trades")%></td>
	        <td width="50%" nowrap><%= l_rs("trarazsoc")%></td>			
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
