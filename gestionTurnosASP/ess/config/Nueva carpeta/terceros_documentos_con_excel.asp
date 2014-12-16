<% Option Explicit 
Dim l_tercero
Dim l_descripcion
l_tercero = request.querystring("desc")'tipotercero
l_descripcion = request.querystring("descripcion")'tercero
%>

<% Response.AddHeader "Content-Disposition", "attachment;filename=" & l_tercero & ".xls" %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->

<% 
'Archivo: terceros_documentos_con_01.asp
'Descripción: Consulta de terceros asociados a tipos de documentos
'Autor : Lisandro Moro
'Fecha: 17/02/2005


Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_tipternro
Dim l_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY tipdocdes "
end if

l_tipternro = request.querystring("tipternro")
l_ternro = request.querystring("cabnro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<head>
<!--<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Documentos asociados al tercero - Ticket</title>
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
		<th colspan="5"><%= l_tercero %><br><%= l_descripcion %></th>
	</tr>
    <tr>
		<th>Sigla</th>
        <th>Descripci&oacute;n</th>
		<th>Número</th>
		<th nowrap>Fecha Vto.</th>
        <th align="center">Habilitado</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql =  " SELECT tkt_tipodocumento.tipdocdes, tipdocsig, oblig, nrodoc, fecvto"
l_sql = l_sql & " from tkt_tipodocumento "
l_sql = l_sql & " INNER JOIN tkt_tipterdoc ON tkt_tipterdoc.tipdocnro = tkt_tipodocumento.tipdocnro "
l_sql = l_sql & " and tkt_tipterdoc.tipternro = " & l_tipternro

l_sql = l_sql & " LEFT JOIN tkt_terdoc ON tkt_terdoc.tipdocnro = tkt_tipodocumento.tipdocnro "
l_sql = l_sql & " and tkt_terdoc.valnro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen Tipos de Documentos asociados</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
			<td width="20%" nowrap><%=l_rs("tipdocsig")%></td>
	        <td width="30%" nowrap><%=l_rs("tipdocdes")%></td>
			<td width="20%" nowrap><%=l_rs("nrodoc")%></td>
			<td width="20%" nowrap><%=l_rs("fecvto")%></td>
			<td width="20%" align="center" nowrap><%if l_rs("oblig") = -1 then %>Si<% Else %>No<% End If %></td>
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
