<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: bancos_con_01.asp
'Descripción: Grilla Administración de bancos
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

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY conceptosCompraVenta.descripcion "
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
<title>Administracion de Articulos</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Articulo</th>
		<th>Compras</th>
		<th>Ventas</th>
		<th>Stock</th>
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "*", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    conceptosCompraVenta.* "	
	l_sql = l_sql & " ,( select SUM( detalleCompras.cantidad) "
	l_sql = l_sql & " from detalleCompras "
	l_sql = l_sql & " where detalleCompras.idconceptoCompraVenta = conceptosCompraVenta.id "
	l_sql = l_sql & " )	as cantidad_compra "
	l_sql = l_sql & " ,( select SUM( detalleVentas.cantidad) "
	l_sql = l_sql & " from detalleVentas "
	l_sql = l_sql & " where detalleVentas.idconceptoCompraVenta = conceptosCompraVenta.id "
	l_sql = l_sql & " )	as cantidad_venta "	
    l_sql = l_sql & " FROM conceptosCompraVenta "
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and conceptosCompraVenta.empnro = " & Session("empnro")   
	else
		l_sql = l_sql & " where conceptosCompraVenta.empnro = " & Session("empnro")   
	end if
	
	
	
    l_sql = l_sql & " " & l_orden

	'response.write l_sql & "</br>"
    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen Articulos cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialog','conceptosCompraVenta_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detalle_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("descripcion")%></td>	
			<td width="10%" align="center" nowrap><%= l_rs("cantidad_compra")%></td>
			<td width="10%" align="center" nowrap><%= l_rs("cantidad_venta")%></td>
			<td width="10%" align="center" nowrap><%= cdbl(l_rs("cantidad_compra")) - cdbl(l_rs("cantidad_venta"))%></td>	
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialog','conceptosCompraVenta_con_02.asp?Tipo=M&cabnro=' + document.detalle_01.cabnro.value,650,250);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detalle_01.cabnro,'dialogAlert','dialogConfirmDelete');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
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
