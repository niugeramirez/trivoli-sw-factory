<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: empresas_con_01.asp
'Descripción: Abm de  Empresas
'Autor : Lisandro Moro
'Fecha: 08/02/2005
' Modificado: 04/01/2006 Raul Chinestra estaba trayendo mal la localidad dado que el join estaba con campo distintos.

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY empcod "
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
<title><%= Session("Titulo")%>Empresas - Ticket</title>
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
        <th align="center">Código</th>
        <th>Descripci&oacute;n</th>		
		<th>Localidad</th>	
		<th>Especial</th>	
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empnro, empcod, empdes, empact, locdes, empesp "
l_sql = l_sql & " FROM tkt_empresa "
l_sql = l_sql & " INNER JOIN tkt_localidad ON tkt_empresa.locnro = tkt_localidad.locnro "
l_sql = l_sql & " WHERE empact = -1 "
if l_filtro <> "" then
  l_sql = l_sql & " AND (" & l_filtro & ") "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Empresas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('empresas_con_02.asp?cabnro=' + datos.cabnro.value,'',700,550)" onclick="Javascript:Seleccionar(this,<%= l_rs("empnro")%>)">
	        <td width="20%" align="center"><%= l_rs("empcod")%></td>
	        <td width="40%" nowrap><%= l_rs("empdes")%></td>
			<td width="35%" nowrap><%= l_rs("locdes")%></td>
			<td width="5%" nowrap><%if l_rs("empesp") = -1 then%>Si<% Else %>No<% End If %></td>			
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
