<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: vendedores_inhabilitados_con_01.asp
'Descripción: Abm de  vendedores_inhabilitados
'Autor : Lisandro Moro
'Fecha: 09/02/2005
on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_todos

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
<script src="/ticket/shared/js/fn_sel_multiple.js"></script>
<head>
<link href="/ticket/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Vendedores Inhabilitados - Ticket</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

/*function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}*/
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>Descripci&oacute;n</th>
		<th>Razón Social</th>
		<th>Tipo</th>
		<th>Habilitado</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT vencornro, vencordes, vencorrazsoc, venhab, vencortip "', nrodoc "
l_sql = l_sql & " FROM tkt_vencor "
l_sql = l_sql & " WHERE venact = -1 AND (vencortip = 'V' OR vencortip = 'C' )"
if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro 
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No existen Vendedores/Corredores Inhabilitados</td>
</tr>
<%else
	l_todos = CStr(trim(l_rs("vencornro")))
	do until l_rs.eof %>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("vencornro")%>)">
	        <td width="20%" nowrap><%= l_rs("vencordes")%></td>
			<td width="80%" nowrap><%= l_rs("vencorrazsoc")%></td>
			<td width="80%" nowrap><% if l_rs("vencortip") = "C" then%>Corredor<% Else %>Vendedor<% End If %></td>
			<td width="40%" nowrap align="center"><% if l_rs("venhab") then %>Si<% Else %>No<% End If %></td>
	    </tr>
	<%	l_rs.MoveNext
		if not l_rs.EoF then
			l_todos = l_todos & "," & Trim(l_rs("vencornro"))
		end if
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
<input type="hidden" name="listanro" value="">
<input type="hidden" name="listatodos" value="<%= l_todos%>">

</form>
<script>
  setearObjDatos(document.datos.listanro, document.datos.listatodos);
</script>

</body>
</html>
