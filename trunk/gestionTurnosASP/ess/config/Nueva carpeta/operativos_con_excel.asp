<% Option Explicit %>
<% Response.AddHeader "Content-Disposition", "attachment;filename=Operativos.xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: operativos_con_excel.asp
'Descripción: Abm de operativos
'Autor : Lisandro Moro
'Fecha: 09/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY opecod "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Carga de Operativos - Ticket</title>
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
		<th colspan="7">Carga de Operativos</th>
	</tr>
    <tr>
        <th align="center" nowrap>Código</th>
		<th nowrap>Cantidad</th>
        <th nowrap>Procedencia</th>
		<th nowrap>Fecha llegada</th>
		<th nowrap>Hora llegada</th>
		<th nowrap>Tipo de Operativo</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT openro, opecod, opecan, lugdes, opefeclle, opehorlle, opetip "
l_sql = l_sql & " FROM tkt_operativo "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_lugar.lugnro = tkt_operativo.lugnro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="7">No existen Operativos</td>
</tr>
<%else
	do until l_rs.eof
	%>
    <tr ondblclick="Javascript:parent.abrirVentana('operativos_con_02.asp?cabnro=' + datos.cabnro.value,'',520,250)" onclick="Javascript:Seleccionar(this,<%= l_rs("openro")%>)">
        <td width="10%" align="center" nowrap><%= l_rs("opecod")%></td>
		<td width="15%" align="center" nowrap><%= l_rs("opecan")%></td>
		<td width="15%" align="center" nowrap><%= l_rs("lugdes")%></td>
		<td width="15%" align="center" nowrap><%= l_rs("opefeclle")%></td>
		<td width="15%" align="center" nowrap><%= left(l_rs("opehorlle"),2) &":"&right(l_rs("opehorlle"),2)%></td>
		<td width="15%" align="center" nowrap><% if l_rs("opetip") = "C" then response.write "Carga" else response.write "Descarga" end if%></td>				
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
