<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: terceros_documentos_con_01.asp
'Descripción: Consulta de documentos asociados a tipos de terceros
'Autor : Lisandro Moro
'Fecha: 17/02/2005

on error GoTo 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_tipternro
Dim l_ternro
Dim l_stringJS

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
<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Documentos asociados al tercero - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,tipter,tipdocnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.tipternro.value = tipter;
	document.datos.tipdocnro.value = tipdocnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
		<th>Sigla</th>
        <th>Descripci&oacute;n</th>
		<th>Número</th>
		<th nowrap>Fecha Vto.</th>
        <th align="center">Obligatorio</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql =  " SELECT tkt_tipodocumento.tipdocdes, tipdocsig, oblig, nrodoc, fecvto, tkt_tipodocumento.tipdocnro"
l_sql = l_sql & " from tkt_tipodocumento "
l_sql = l_sql & " INNER JOIN tkt_tipterdoc ON tkt_tipterdoc.tipdocnro = tkt_tipodocumento.tipdocnro "
l_sql = l_sql & " and tkt_tipterdoc.tipternro = " & l_tipternro
l_sql = l_sql & " LEFT JOIN tkt_terdoc ON tkt_terdoc.tipdocnro = tkt_tipodocumento.tipdocnro "
l_sql = l_sql & " and tkt_terdoc.valnro = " & l_ternro
l_sql = l_sql & " and tkt_terdoc.tipternro = " & l_tipternro

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="5">No existen ningún Documento asociado</td>
</tr>
<%else			
	'tipternro	tipterdes
	'1			Vendedor
	'2			Camionero
	'3			Entregador/Recibidor
	'4			Empresa
	'5			Destinatario
	'6			Transportista
	'7			Corredor
	'8			Vendedor/Corredor
	'9			Entregador
	'10			Recibidor
	'11			Cuenta y Orden
  	if ((l_tipternro = 3) OR (l_tipternro = 9) OR (l_tipternro = 10)) Then
		l_StringJS = "Javascript:parent.abrirVentana('terceros_documentos_con_02.asp?tipternro=" & l_tipternro & "&ternro=' + document.datos.cabnro.value + '&tipdocnro=' + document.datos.tipdocnro.value,'',400,135)"
	else
		l_StringJS = ""
	end if
	do until l_rs.eof
	%>
	    <tr ondblclick="<%= l_StringJS %>" onclick="Javascript:Seleccionar(this,<%= l_ternro%>,<%= l_tipternro %>,<%= l_rs("tipdocnro") %>)">
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
<input type="hidden" name="tipternro" value="0">
<input type="hidden" name="tipdocnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
