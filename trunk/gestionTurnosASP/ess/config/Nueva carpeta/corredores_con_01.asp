<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: corredores_con_01.asp
'Descripción: Abm de corredores
'Autor : Gustavo Manfrin
'Fecha: 09/02/2005
'Modificado: 

on error goto 0

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden


l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY vencordes "
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
<title><%= Session("Titulo")%>corredores - Ticket</title>
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
        <th align="center">C&oacute;digo</th>
        <th>Nombre clave</th>		
        <th>Razón social</th>		
        <th>Hab.</th>				
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT vencornro, vencorcod, vencordes, vencorrazsoc, venhab "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND vencortip = 'C' "
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 

if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen corredores</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('corredores_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',550,170)" onclick="Javascript:Seleccionar(this,<%= l_rs("vencornro")%>)">
	        <td width="15%" align="center"><%= l_rs("vencorcod")%></td>
	        <td width="25%" nowrap><%= l_rs("vencordes")%></td>
	        <td width="50%" nowrap><%= l_rs("vencorrazsoc")%></td>			
            <td width="10%" align="center"><% if l_rs("venhab") = -1 then%> Si <% Else %> No <% End If%></td>			
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
