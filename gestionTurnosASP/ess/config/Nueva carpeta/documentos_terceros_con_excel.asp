<% Option Explicit 
Dim l_documento
l_documento = request.querystring("desc")
%>

<% Response.AddHeader "Content-Disposition", "attachment;filename=" & l_documento & ".xls" %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
'Archivo: documentos_terceros_con_01.asp
'Descripción: Consulta de terceros asociados a tipos de documentos
'Autor : Lisandro Moro
'Fecha: 17/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

Dim l_tipdocnro


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY tipterdes "
end if

l_tipdocnro = request.querystring("cabnro")


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Terceros asociados al documento - Ticket</title>
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
		<th  colspan="2"><%= l_documento %></th>
	</tr>
    <tr>
        <th>Descripci&oacute;n</th>		
        <th align="center">Estado</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql =  " SELECT tkt_tipotercero.tipterdes, oblig"
l_sql = l_sql & " from tkt_tipotercero "
l_sql = l_sql & " INNER JOIN tkt_tipterdoc ON tkt_tipterdoc.tipternro = tkt_tipotercero.tipternro "
l_sql = l_sql & " and tkt_tipterdoc.tipdocnro= " & l_tipdocnro
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Tipos de Terceros asociados</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr>
	        <td width="80%" nowrap><%=l_rs("tipterdes")%></td>
			<td width="20%" align="center"><%if l_rs("oblig") = -1 then %>Obligatorio<% Else %>No Obligatorio<% End If %></td>
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
