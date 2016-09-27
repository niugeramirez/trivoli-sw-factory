<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: ventas_con_01.asp
'Descripción: Grilla Administración de Ventas
'Autor : Trivoli
'Fecha: 31/05/2015

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_saldo
Dim l_cobrado
Dim l_monto_venta
Dim l_cant

Dim l_primero

l_filtro = request("filtro")
'response.write l_filtro 
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY ventas.fecha desc "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administracion de Ventas</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Nombre</th>
        <th>Fecha</th>	
		<th>Monto</th>
		<th>Costo</th>
		<th>Utilidad</th>
		<th>Cobrado</th>	
		<th>Saldo</th>
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "**", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    ventas.* , clientes.nombre "
	l_sql = l_sql & " ,( select SUM( detalleVentas.cantidad*detalleVentas.precio_unitario) "
	l_sql = l_sql & " 	from detalleVentas "
	l_sql = l_sql & " 	where detalleVentas.idVenta = ventas.id "
	l_sql = l_sql & " )	as monto_venta "
	l_sql = l_sql & " ,(	select SUM( costosVentas.cantidad*costosVentas.precio_unitario) "
	l_sql = l_sql & " 	from costosVentas "
	l_sql = l_sql & " 	where costosVentas.idVenta = ventas.id "
	l_sql = l_sql & " )	as costo_venta "
	l_sql = l_sql & " ,(	select	SUM(cajaMovimientos.monto) "
	l_sql = l_sql & " 	from	cajaMovimientos "
	l_sql = l_sql & " 	where	cajaMovimientos.idventaOrigen = ventas.id "
	l_sql = l_sql & " )	as  cobrado	"
    l_sql = l_sql & " FROM ventas "
    l_sql = l_sql & " LEFT JOIN clientes ON clientes.id = ventas.idcliente "
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and ventas.empnro = " & Session("empnro")   
	else
		l_sql = l_sql & " where ventas.empnro = " & Session("empnro")   
	end if
	
	
	
    l_sql = l_sql & " " & l_orden

	'response.write l_sql & "</br>"
    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen Ventas cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','ventas_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("nombre")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("fecha")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("monto_venta")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("costo_venta")%></td>
			<td width="10%" align="center" nowrap><%= round(cdbl(l_rs("monto_venta")) - cdbl(l_rs("costo_venta")),2)%></td>
			<td width="10%" align="center" nowrap><%= l_rs("cobrado")%></td>
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
			<td width="10%" align="center" nowrap><%= Round(l_saldo,2) %></td>			
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialog','ventas_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,250);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01.cabnro,'dialogAlert','dialogConfirmDelete');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
											 
				<a href="Javascript:parent.abrirDialogo('dialog_cont_DV','detalleventa_con_00.asp?p_carga_js=N&id=' + document.detalle_01.cabnro.value,620,400);"><img src="../shared/images/Data-List-icon_16.png" border="0" title="Detalle de Venta"></a>
				<a href="Javascript:parent.abrirDialogo('dialog_cont_CV','costoventa_con_00.asp?p_carga_js=N&id=' + document.detalle_01.cabnro.value,620,400);"><img src="../shared/images/Ecommerce-Price-Tag-icon.png" border="0" title="Costo de Venta"></a>
				<a href="Javascript:parent.abrirDialogo('dialog_cont_CM','cajamovimientos_con_00.asp?p_carga_js=N&p_id_venta=' + document.detalle_01.cabnro.value,1020,500);"><img src="../shared/images/US-dollar-icon_16.png" border="0" title="Pagos"></a>
			</td>
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
<form name="detalle_01" id="detalle_01" method="post">
    <input type="hidden" id="cabnro" name="cabnro" value="0">
    <input type="hidden" name="orden" value="<%= l_orden %>">
    <input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
