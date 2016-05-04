<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: cajamovimientos_con_01.asp
'Descripción: Grilla Administración de cajamovimientos
'Autor : Trivoli
'Fecha: 31/05/2015

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant

Dim l_primero
Dim l_p_id_venta
Dim l_p_id_compra




l_filtro = request("filtro")
l_orden  = request("orden")

l_p_id_venta = request.querystring("p_id_venta")
l_p_id_compra = request.querystring("p_id_compra")
'response.write  "p_id_compra "&l_p_id_compra&"</br>"

if l_orden = "" then
  l_orden = " ORDER BY cajamovimientos.fecha desc "
end if
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script>
	function editar_registro(){
		parent.abrirDialogo(
							'dialog_mc'
							,'cajamovimientos_con_02.asp?Tipo=M&cabnro=' + document.detalle_01_mc.cabnro.value+'&p_id_venta='+'<%= l_p_id_venta%>' +'&p_id_compra='+'<%= l_p_id_compra%>' 
							,650
							,450
							);
	}	
</script>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Administracion de Caja</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar_mc();">
<table>
    <tr>
        <th>Fecha</th>		
        <th>Tipo</th>			
        <th>Movimiento</th>			
        <th>Detalle</th>		
		<th>Unidad de Negocio</th>
		<th>Medio de Pago</th>
		<th>Cheque</th>
		<th>Monto</th>
		<th>Responsable</th>
		<th>Operacion</th>
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "*", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    cajamovimientos.* , tiposmovimientocaja.descripcion , unidadesNegocio.descripcion unidadnegocio , mediosdepago.titulo , bancos.nombre_banco, cheques.numero, responsablesCaja.nombre responsable "
	l_sql = l_sql & " ,compras.fecha as fecha_compra "
	l_sql = l_sql & " ,proveedores.nombre as nombre_proveedor "
	l_sql = l_sql & " ,ventas.fecha as fecha_venta "
	l_sql = l_sql & " ,clientes.nombre as nombre_cliente "    
	l_sql = l_sql & " FROM cajamovimientos "
    l_sql = l_sql & " LEFT JOIN tiposmovimientocaja ON tiposmovimientocaja.id = cajamovimientos.idtipoMovimiento "
    l_sql = l_sql & " LEFT JOIN unidadesNegocio ON unidadesNegocio.id = cajamovimientos.idunidadnegocio "
    l_sql = l_sql & " LEFT JOIN mediosdepago ON mediosdepago.id = cajamovimientos.idmediopago "	
    l_sql = l_sql & " LEFT JOIN cheques ON cheques.id = cajamovimientos.idcheque "		
	l_sql = l_sql & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "
    l_sql = l_sql & " LEFT JOIN responsablesCaja ON responsablesCaja.id = cajamovimientos.idresponsable "	
    l_sql = l_sql & " LEFT JOIN compras ON compras.id = cajaMovimientos.idcompraOrigen "
    l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "
    l_sql = l_sql & " LEFT JOIN ventas ON ventas.id = cajaMovimientos.idventaOrigen "
    l_sql = l_sql & " LEFT JOIN clientes ON clientes.id = ventas.idcliente "	
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and cajamovimientos.empnro = " & Session("empnro")   
	else
		l_sql = l_sql & " where cajamovimientos.empnro = " & Session("empnro")   
	end if
	
	if l_p_id_venta <> "" then 
		l_sql = l_sql & " and cajaMovimientos.idventaOrigen = " & l_p_id_venta
	end if

	if l_p_id_compra <> "" then 
		l_sql = l_sql & " and cajaMovimientos.idcompraOrigen = " & l_p_id_compra
	end if	
	
    l_sql = l_sql & " " & l_orden

	'response.write l_sql & "</br>"
    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen Movimientos de Caja cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:editar_registro();" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01_mc.cabnro)">    
			<td width="10%" nowrap><%= l_rs("fecha")%></td>
			
			<td width="10%" align="left" nowrap><% if l_rs("tipo") = "E" then response.write "Entrada" else response.write  "Salida" end if%></td>			
			<td width="10%" align="left" nowrap><%= l_rs("descripcion")%></td>
			<td width="10%" align="left" ><%= l_rs("detalle")%></td>			
			
			<td width="10%" align="left" nowrap><%= l_rs("unidadnegocio")%></td>	
				
	        <td width="10%" nowrap align="center"><%= l_rs("titulo")%></td>			
			
			<td width="10%" nowrap align="center"><%= l_rs("nombre_banco") %> &nbsp;-&nbsp;<%= l_rs("numero") %></td>	
			<td width="10%" nowrap align="left"><%= l_rs("monto")%></td>		
			<td width="10%" nowrap align="left"><%= l_rs("responsable")%></td>

			<td width="10%" nowrap align="left">
				<%= l_rs("nombre_cliente")%> </br> <%= l_rs("fecha_venta")%>
				<%= l_rs("nombre_proveedor")%> </br> <%= l_rs("fecha_compra")%>
			</td>			
			  	
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:editar_registro();"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01_mc.cabnro,'dialogAlert_mc','dialogConfirmDelete_mc');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
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
<form name="detalle_01_mc" id="detalle_01_mc" method="post">
    <input type="hidden" id="cabnro" name="cabnro" value="0">
    <input type="hidden" name="orden" value="<%= l_orden %>">
    <input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
