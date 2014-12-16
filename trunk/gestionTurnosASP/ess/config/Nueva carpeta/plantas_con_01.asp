<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: plantas_con_01.asp
'Descripción: Abm de Plantas de Trabajo
'Autor : Alvaro Bayon
'Fecha: 08/02/2005
'Modificado: 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY plades "
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
<title><%= Session("Titulo")%>Plantas - Ticket</title>
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
        <th>Descripci&oacute;n</th>
        <th>Localidad</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT planro, plades, locdes "
l_sql = l_sql & " FROM tkt_planta "
l_sql = l_sql & " INNER JOIN tkt_localidad ON  tkt_localidad.locnro = tkt_planta.locnro"
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Plantas</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('plantas_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,140)" onclick="Javascript:Seleccionar(this,<%= l_rs("planro")%>)">
	        <td width="50%" nowrap><%= l_rs("plades")%></td>
	        <td width="50%" nowrap><%= l_rs("locdes")%></td>
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
