<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_01.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")


if l_orden = "" then
  l_orden = " ORDER BY  nombre "
end if


'l_ternro  = request("ternro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Buscar Proveedores</title>
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Apellido y Nombre</th>
        <th>Tel</th>		
        <th>Celular</th>		
        <th>Direccion</th>		
					
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    proveedores.id id, proveedores.*  "
l_sql = l_sql & " FROM proveedores "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql  = l_sql  & " AND proveedores.empnro = " & Session("empnro")
end if

if l_filtro = "" then
  l_sql  = l_sql  & " WHERE proveedores.empnro = " & Session("empnro")
end if

l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Proveedores cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr ondblclick="Javascript:parent.AsignarProveedor('<%= l_rs("id")%>','<%= l_rs("nombre")%>' )" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">

	        <td width="10%" nowrap><%= l_rs("nombre")%></td>
	        <td width="10%" nowrap><%= l_rs("telefono")%></td>		
	        <td width="10%" align="center"><%= l_rs("celular")%></td>								
	        <td width="10%" nowrap><%= l_rs("direccion")%></td>			
		

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
