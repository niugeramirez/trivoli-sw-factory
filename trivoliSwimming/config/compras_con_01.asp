<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: compras_con_01.asp
'Descripción: Grilla Administración de Compras
'Autor : Trivoli
'Fecha: 31/05/2015

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_saldo
Dim l_pagado
Dim l_monto_compra


Dim l_cant

Dim l_primero

l_filtro = request("filtro")
'response.write l_filtro &"</br>"&"</br>"
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY compras.fecha desc "
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
<title>Administracion de Compras</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Nombre</th>
        <th>Fecha</th>	
		<th>Monto</th>
		<th>Pagado</th>	
		<th>Saldo</th>		
		<th>Acciones</th>		
    </tr>
    <%
    
	l_filtro = replace (l_filtro, "**", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    compras.* , proveedores.nombre "	
    l_sql = l_sql & " ,( select SUM( detalleCompras.cantidad*detalleCompras.precio_unitario) "
    l_sql = l_sql & " 	from detalleCompras "
	l_sql = l_sql & " where detalleCompras.idcompra = compras.id "
    l_sql = l_sql & " )	as monto_compra "
    l_sql = l_sql & " ,(	select	SUM(cajaMovimientos.monto) "
	l_sql = l_sql & " from	cajaMovimientos "
	l_sql = l_sql & " where	cajaMovimientos.idcompraOrigen = compras.id "
    l_sql = l_sql & " )	as  pagado "	
    l_sql = l_sql & " FROM compras "
    l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "
	' Multiempresa
	l_sql = l_sql & " where compras.empnro = " & Session("empnro") 	
	
	if l_filtro <> "" then	  
	  l_sql = l_sql & " and " & l_filtro & " "   
	end if	
	
    l_sql = l_sql & " " & l_orden
	
	'response.write l_sql & "</br>"
    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen Compras cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','compras_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("nombre")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("fecha")%></td>			
			<td width="10%" align="center" nowrap><%= l_rs("monto_compra")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("pagado")%></td>			
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
			<td width="10%" align="center" nowrap><%= round(l_saldo,2) %></td>			
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialog','compras_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,250);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01.cabnro,'dialogAlert','dialogConfirmDelete');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
				
				<a href="Javascript:parent.abrirDialogo('dialog_cont_DC','detallecompra_con_00.asp?p_carga_js=N&id=' + document.detalle_01.cabnro.value,620,400);"><img src="../shared/images/Data-List-icon_16.png" border="0" title="Detalle de Compra"></a>
				<a href="Javascript:parent.abrirDialogo('dialog_cont_CMC','cajamovimientos_con_00.asp?p_carga_js=N&p_id_compra=' + document.detalle_01.cabnro.value,1020,500);"><img src="../shared/images/US-dollar-icon_16.png" border="0" title="Pagos"></a>				
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
