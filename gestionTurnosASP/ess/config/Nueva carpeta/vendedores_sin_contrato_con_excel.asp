<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Vendedores habilitados a descargar sin contrato.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: vendedores_sin_contrato_con_excel.asp
'Descripción: Abm de vendedores habilitados a descatrgar sin_contrato
'Autor : Lisandro Moro
'Fecha: 10/02/2005

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
  l_orden = " ORDER BY vencordes "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores Habilitados a Descargar sin Contrato - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,pronro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.pronro.value = pronro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
		<th colspan="4">Vendedores habilitados a descargar sin contrato</th>
	</tr>
    <tr>
        <th>Descripci&oacute;n</th>
		<th>Razón Social</th>
		<th>Producto</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT tkt_vencor.vencornro, vencordes, vencorrazsoc, prodes, vencortip, tkt_producto.pronro "', nrodoc "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " LEFT JOIN tkt_autsincon ON tkt_vencor.vencornro = tkt_autsincon.vencornro "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_producto.pronro = tkt_autsincon.pronro "
l_sql = l_sql & " WHERE venact = -1 AND venhab = -1 and tkt_vencor.vencortip='V' "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Vendedores Habilitados a Descargar sin Contrato</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('vendedores_sin_contrato_con_02.asp?tipo=M&cabnro=' + datos.cabnro.value + '&pronro=' + datos.pronro.value,'',520,160)" onclick="Javascript:Seleccionar(this,<%= l_rs("vencornro")%>, <%= l_rs("pronro") %>)">
	        <td width="20%" nowrap><%= l_rs("vencordes")%></td>
			<td width="80%" nowrap><%= l_rs("vencorrazsoc")%></td>
			<td width="40%" nowrap><%= l_rs("prodes")%></td>
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
<input type="hidden" name="pronro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
