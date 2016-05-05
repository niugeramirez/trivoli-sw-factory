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
Dim l_saldo
Dim l_pagado
Dim l_monto_compra

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
<title>Buscar Pacientes</title>
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
        <th>Proveedor</th>
        <th>Fecha</th>		
		<th>Saldo</th>	
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    compras.id, compras.fecha, proveedores.nombre  "
l_sql = l_sql & " ,( select SUM( detalleCompras.cantidad*detalleCompras.precio_unitario) "
l_sql = l_sql & " 	from detalleCompras "
l_sql = l_sql & " where detalleCompras.idcompra = compras.id "
l_sql = l_sql & " )	as monto_compra "
l_sql = l_sql & " ,(	select	SUM(cajaMovimientos.monto) "
l_sql = l_sql & " from	cajaMovimientos "
l_sql = l_sql & " where	cajaMovimientos.idcompraOrigen = compras.id "
l_sql = l_sql & " )	as  pagado "	
l_sql = l_sql & " FROM compras "
l_sql = l_sql & " INNER JOIN proveedores ON proveedores.id = compras.idproveedor "


if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql  = l_sql  & " AND compras.empnro = " & Session("empnro")
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
	    <tr ondblclick="Javascript:parent.AsignarCompraOrigen('<%= l_rs("id")%>','<%= l_rs("fecha")%>','<%= l_rs("nombre")%>' )" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">

	        <td width="10%" nowrap><%= l_rs("nombre")%></td>
	        <td align="center" width="10%" nowrap><%= l_rs("fecha")%></td>		
			<% 
				if isnull(l_rs("monto_compra")) then					 
					l_monto_compra = 0
				else					
					l_monto_compra = cdbl(l_rs("monto_compra"))
				end if

				if isnull(l_rs("pagado")) then
					l_pagado = 0
				else 
					l_pagado = cdbl(l_rs("pagado"))
				end if
				
				l_saldo = l_monto_compra - l_pagado 
				if l_saldo = 0 then
					l_saldo = ""
				end if
			%>
			<td width="10%" align="center" nowrap><%= l_saldo %></td>		

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
