<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: proyeccionventas_con_01.asp
'Descripción: Grilla Administración de proyeccionventas
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
  l_orden = " ORDER BY proyeccionventas.fecha_desde "
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
<title>Proyeccion de Ventas</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Fecha Desde</th>
		<th>Fecha Hasta</th>
		<th>Articulo</th>
		<th>Cantidad Proyectada</th>
		<th>Ventas Reales</th>
		<th>Acciones</th>		
    </tr>
    <%
    l_filtro = replace (l_filtro, "*", "%")

    Set l_rs = Server.CreateObject("ADODB.RecordSet")
    l_sql = "SELECT    proyeccionventas.* ,conceptosCompraVenta.descripcion as articulo_desc "
	l_sql = l_sql & " ,(	select sum(detalleVentas.cantidad) "
	l_sql = l_sql & " 	from detalleVentas "
	l_sql = l_sql & " 	inner join ventas on ventas.id = detalleVentas.idventa "
	l_sql = l_sql & " 	where detalleVentas.idconceptoCompraVenta = proyeccionventas.idconceptoCompraVenta "
	l_sql = l_sql & " 	and ventas.fecha >= proyeccionventas.fecha_desde "
	l_sql = l_sql & " 	and ventas.fecha <= proyeccionventas.fecha_hasta "
	l_sql = l_sql & " )									as cant_vtas_reales "		
    l_sql = l_sql & " FROM proyeccionventas "
	l_sql = l_sql & " inner join conceptosCompraVenta on conceptosCompraVenta.id = proyeccionventas.idconceptoCompraVenta "
	' Multiempresa
	if l_filtro <> "" then
	  l_sql = l_sql & " WHERE " & l_filtro & " "
	  l_sql = l_sql & " and proyeccionventas.empnro = " & Session("empnro")   
	else
		l_sql = l_sql & " where proyeccionventas.empnro = " & Session("empnro")   
	end if
	
	
	
    l_sql = l_sql & " " & l_orden

	'response.write l_sql
	
    rsOpen l_rs, cn, l_sql, 0 
    if l_rs.eof then
	    l_primero = 0
    %>
    
    <tr>
	    <td colspan="4" >No existen proyecciones cargados para el filtro ingresado.</td>
    </tr>
    <%
    else
        l_primero = l_rs("id")
	    l_cant = 0
	    do until l_rs.eof
		    l_cant = l_cant + 1
	    %>
	    <tr ondblclick="Javascript:parent.abrirDialogo('dialogPV','proyeccionventas_con_02.asp?Tipo=M&cabnro=' + document.detallePV_01.cabnro.value,650,350);" onclick="Javascript:parent.Seleccionar(this,<%= l_rs("id")%>,document.detallePV_01.cabnro)">    
			<td width="10%" nowrap><%= l_rs("fecha_desde")%></td>	
			<td width="10%" nowrap><%= l_rs("fecha_hasta")%></td>
			<td width="10%" nowrap><%= l_rs("articulo_desc")%></td>		
			<td width="10%" nowrap><%= l_rs("cantidadproyectada")%></td>		
			<td width="10%" nowrap><%= l_rs("cant_vtas_reales")%></td>		
	        <td align="center" width="10%" nowrap>                    
                <a href="Javascript:parent.abrirDialogo('dialogPV','proyeccionventas_con_02.asp?Tipo=M&cabnro=' + document.detallePV_01.cabnro.value,650,250);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>				                																												
				<a href="Javascript:parent.eliminarRegistroAJAX(document.detallePV_01.cabnro,'dialogAlertPV','dialogConfirmDeletePV');"><img src="../shared/images/Eliminar_16.png" border="0" title="Baja"></a>
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
<form name="detallePV_01" id="detallePV_01" method="post">
    <input type="hidden" id="cabnro" name="cabnro" value="0">
    <input type="hidden" name="orden" value="<%= l_orden %>">
    <input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
