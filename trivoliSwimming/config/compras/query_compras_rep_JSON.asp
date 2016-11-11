<% Option Explicit %>

<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<% 





Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_sql2
Dim l_rta
Dim l_data_master
Dim l_data_detail
Dim l_data_detail_pagos

Dim l_nro_cheque
'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")


l_sql = "SELECT    compras.* , proveedores.nombre as nombre_proveedor "	
l_sql = l_sql & " ,( select SUM( detalleCompras.cantidad*detalleCompras.precio_unitario) "
l_sql = l_sql & " 	from detalleCompras "
l_sql = l_sql & " where detalleCompras.idcompra = compras.id "
l_sql = l_sql & " )	as monto_compra "
l_sql = l_sql & " ,(	select	SUM(cajaMovimientos.monto) "
l_sql = l_sql & " from	cajaMovimientos "
l_sql = l_sql & " where	cajaMovimientos.idcompraOrigen = compras.id "
l_sql = l_sql & " )	as  pagado "	
l_sql = l_sql & " ,( select SUM( detalleCompras.cantidad*detalleCompras.precio_unitario) "
l_sql = l_sql & " 	from detalleCompras "
l_sql = l_sql & " where detalleCompras.idcompra = compras.id "
l_sql = l_sql & " )	"
l_sql = l_sql & " - (	select	SUM(cajaMovimientos.monto) "
l_sql = l_sql & " from	cajaMovimientos "
l_sql = l_sql & " where	cajaMovimientos.idcompraOrigen = compras.id "
l_sql = l_sql & " )	as  saldo "	
l_sql = l_sql & " FROM compras "
l_sql = l_sql & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "
' Multiempresa
l_sql = l_sql & " where compras.empnro = " & Session("empnro") 	
l_sql = l_sql & " and proveedores.nombre like 'iGUi%' "
l_sql = l_sql & " and compras.id < 50 "
l_sql = l_sql & " order by compras.fecha desc, proveedores.nombre "

'******************************************ARMO EL JSON DEL DATA MASTER***********************************************
l_data_master = "[ "
l_data_detail = "[ "
l_data_detail_pagos = "[ "

rsOpen l_rs, cn, l_sql, 0
do until l_rs.eof
	l_data_master = l_data_master & "{"
	
	l_data_master = l_data_master & """id_compra"":""" & l_rs("id") 
	l_data_master = l_data_master & """,""nombre_proveedor"":""" & l_rs("nombre_proveedor") & """"	
	l_data_master = l_data_master & ",""fecha"":""" & l_rs("fecha") & """"
	l_data_master = l_data_master & ",""monto_compra"":""" & l_rs("monto_compra") & """"
	l_data_master = l_data_master & ",""pagado"":""" & l_rs("pagado") & """"
	l_data_master = l_data_master & ",""saldo"":""" & l_rs("saldo") & """"
	
	l_data_master = l_data_master &"},"
	
	'*********************************ARMO EL JSON DEL DETALLE*****************************************************
	l_sql2 = "SELECT    detallecompras.*  , conceptosCompraVenta.descripcion as descripcion_articulo"
	l_sql2 = l_sql2 & " ,(detallecompras.cantidad * detallecompras.precio_unitario) as subtotal "
    l_sql2 = l_sql2 & " FROM detallecompras "
    l_sql2 = l_sql2 & " LEFT JOIN conceptosCompraVenta ON conceptosCompraVenta.id = detallecompras.idconceptoCompraVenta "
	l_sql2 = l_sql2 & " where detallecompras.idcompra = " & l_rs("id")

	rsOpen l_rs2, cn, l_sql2, 0
	do until l_rs2.eof
		l_data_detail = l_data_detail & "{"
		
		l_data_detail = l_data_detail & """descripcion_articulo"":""" & l_rs2("descripcion_articulo") 
		l_data_detail = l_data_detail & """,""cantidad"":""" & l_rs2("cantidad") & """"	
		l_data_detail = l_data_detail & ",""precio_unitario"":""" & l_rs2("precio_unitario") & """"
		l_data_detail = l_data_detail & ",""subtotal"":""" & l_rs2("subtotal") & """"
		l_data_detail = l_data_detail & ",""id_compra"":""" & l_rs2("idcompra") & """"

		
		l_data_detail = l_data_detail &"},"
	
		l_rs2.MoveNext
	loop
	l_rs2.Close	
	'*********************************ARMO EL JSON DEL DETALLE 	DE PAGOS********************************************
    l_sql2 = "SELECT    cajamovimientos.* , tiposmovimientocaja.descripcion , unidadesNegocio.descripcion unidadnegocio , mediosdepago.titulo as medio_pago "
	l_sql2 = l_sql2 & " , bancos.nombre_banco " 	
	l_sql2 = l_sql2 & " , cheques.numero as numero "
	l_sql2 = l_sql2 & " , responsablesCaja.nombre responsable "
	l_sql2 = l_sql2 & " ,compras.fecha as fecha_compra "
	l_sql2 = l_sql2 & " ,proveedores.nombre as nombre_proveedor "
	l_sql2 = l_sql2 & " ,ventas.fecha as fecha_venta "
	l_sql2 = l_sql2 & " ,clientes.nombre as nombre_cliente "    
	l_sql2 = l_sql2 & " FROM cajamovimientos "
    l_sql2 = l_sql2 & " LEFT JOIN tiposmovimientocaja ON tiposmovimientocaja.id = cajamovimientos.idtipoMovimiento "
    l_sql2 = l_sql2 & " LEFT JOIN unidadesNegocio ON unidadesNegocio.id = cajamovimientos.idunidadnegocio "
    l_sql2 = l_sql2 & " LEFT JOIN mediosdepago ON mediosdepago.id = cajamovimientos.idmediopago "	
    l_sql2 = l_sql2 & " LEFT JOIN cheques ON cheques.id = cajamovimientos.idcheque "		
	l_sql2 = l_sql2 & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "
    l_sql2 = l_sql2 & " LEFT JOIN responsablesCaja ON responsablesCaja.id = cajamovimientos.idresponsable "	
    l_sql2 = l_sql2 & " LEFT JOIN compras ON compras.id = cajaMovimientos.idcompraOrigen "
    l_sql2 = l_sql2 & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor "
    l_sql2 = l_sql2 & " LEFT JOIN ventas ON ventas.id = cajaMovimientos.idventaOrigen "
    l_sql2 = l_sql2 & " LEFT JOIN clientes ON clientes.id = ventas.idcliente "					
	l_sql2 = l_sql2 & " WHERE cajaMovimientos.idcompraOrigen = " & l_rs("id")	
    l_sql2 = l_sql2 & " ORDER BY cajaMovimientos.fecha" 	
	
	rsOpen l_rs2, cn, l_sql2, 0
	do until l_rs2.eof
		l_data_detail_pagos = l_data_detail_pagos & "{"
		

		
		l_data_detail_pagos = l_data_detail_pagos & """fecha"":""" & l_rs2("fecha") 
		l_data_detail_pagos = l_data_detail_pagos & """,""medio_pago"":""" & l_rs2("medio_pago") & """"					
		l_data_detail_pagos = l_data_detail_pagos & ",""nombre_banco"":""" & l_rs2("nombre_banco") & """"			
		l_data_detail_pagos = l_data_detail_pagos & ",""cheque"":""" & l_rs2("numero") & """"			
		l_data_detail_pagos = l_data_detail_pagos & ",""monto"":""" & l_rs2("monto") & """"
		l_data_detail_pagos = l_data_detail_pagos & ",""id_compra"":""" & l_rs2("idcompraOrigen") & """"	
		
		l_data_detail_pagos = l_data_detail_pagos &"},"
	
		l_rs2.MoveNext
	loop
	l_rs2.Close		
	'**************************************************************************************************************
	
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing

l_data_master=Left(l_data_master,Len(l_data_master)-1)
l_data_master = l_data_master & "]"			

l_data_detail=Left(l_data_detail,Len(l_data_detail)-1)
l_data_detail = l_data_detail & "]"

l_data_detail_pagos=Left(l_data_detail_pagos,Len(l_data_detail_pagos)-1)
l_data_detail_pagos = l_data_detail_pagos & "]"
'*********************************************************************************************************************
 	


l_rta = " [ {""data_master"": "&l_data_master&"}"
l_rta = l_rta&",  {""data_detail"": "&l_data_detail&"}"
l_rta = l_rta&",  {""data_detail_pagos"": "&l_data_detail_pagos&"}"
l_rta = l_rta& "]"

Response.write l_rta

%>

