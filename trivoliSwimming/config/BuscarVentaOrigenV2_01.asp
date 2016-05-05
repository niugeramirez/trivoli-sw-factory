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
Dim l_cobrado
Dim l_monto_venta
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")


if l_orden = "" then
  l_orden = " ORDER BY  fecha desc "
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
        <th>Cliente</th>
        <th>Fecha</th>		
		<th>Saldo</th>
					
    </tr>
<%
l_filtro = replace (l_filtro, "**", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    ventas.id, ventas.fecha, clientes.nombre  "
l_sql = l_sql & " ,( select SUM( detalleVentas.cantidad*detalleVentas.precio_unitario) "
l_sql = l_sql & " 	from detalleVentas "
l_sql = l_sql & " 	where detalleVentas.idVenta = ventas.id "
l_sql = l_sql & " )	as monto_venta "
l_sql = l_sql & " ,(	select	SUM(cajaMovimientos.monto) "
l_sql = l_sql & " 	from	cajaMovimientos "
l_sql = l_sql & " 	where	cajaMovimientos.idventaOrigen = ventas.id "
l_sql = l_sql & " )	as  cobrado	"	
l_sql = l_sql & " FROM ventas "
l_sql = l_sql & " INNER JOIN clientes ON clientes.id = ventas.idcliente "


if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql  = l_sql  & " AND ventas.empnro = " & Session("empnro")
else  
  l_sql = l_sql & " where ventas.empnro = " & Session("empnro")   
end if



l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Clientes cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr ondblclick="Javascript:parent.AsignarVentaOrigen('<%= l_rs("id")%>','<%= l_rs("fecha")%>','<%= l_rs("nombre")%>' )" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">

	        <td width="10%" nowrap><%= l_rs("nombre")%></td>
	        <td align="center" width="10%" nowrap><%= l_rs("fecha")%></td>		
			<% 
				if isnull(l_rs("monto_venta")) then					 
					l_monto_venta = 0
				else					
					l_monto_venta = cdbl(l_rs("monto_venta"))
				end if

				if isnull(l_rs("cobrado")) then
					l_cobrado = 0
				else 
					l_cobrado = cdbl(l_rs("cobrado"))
				end if
				
				l_saldo = l_monto_venta - l_cobrado 
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
