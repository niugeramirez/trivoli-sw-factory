<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_numero
Dim l_idcompraorigen
Dim l_idventaorigen
dim l_idtipomovimiento

Dim l_monto
Dim l_aux


Dim texto

texto = ""
l_tipo		    = request.Form("tipo")
l_id            = request.Form("id")
l_numero	 	= request.Form("numero")
l_idcompraorigen = request.Form("idcompraorigen")
l_idventaorigen = request.Form("idventaorigen")
l_idtipomovimiento = request.Form("idtipomovimiento") 
l_monto = request.Form("monto")

if len(l_idventaorigen) = 0 then
	l_idventaorigen = "0"
end if

if len(l_idcompraorigen) = 0 then
	l_idcompraorigen = "0"
end if

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")


texto = "OK"

'response.write "l_idventaorigen "&l_idventaorigen&" l_idcompraorigen "&l_idcompraorigen&" </br>"

'No se puede asociar a una compra y a  una venta simultaneamente
if texto = "OK" then
	if l_idcompraorigen <> "0" and l_idventaorigen <> "0" then
		texto =  "No puede asociar el mismo movimiento a una compra y a una venta."&"</br>"&"</br>"
		texto = texto & "Selecciones una compra o una venta. No ambas."
	end if
end if

'Chequeo que si el tipo de movimiento es venta se complete una venta
if texto = "OK" then

	l_sql = "SELECT * "
	l_sql = l_sql & " FROM tiposMovimientoCaja "
	l_sql = l_sql & " WHERE tiposMovimientoCaja.id ='" & l_idtipomovimiento & "'"
	l_sql = l_sql & " and tiposMovimientoCaja.flagVenta = -1 " 
	
	l_sql = l_sql & " and tiposMovimientoCaja.empnro = " & Session("empnro")   

	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof and l_idventaorigen = "0" then
	    texto =  "El tipo de movimiento es "& l_rs("descripcion") &".</br>"&"</br>"
		texto =  texto& "Debe seleccionar una Venta."
	end if 
	l_rs.close
end if

'Chequeo que si el tipo de movimiento es compra se complete una compra
if texto = "OK" then

	l_sql = "SELECT * "
	l_sql = l_sql & " FROM tiposMovimientoCaja "
	l_sql = l_sql & " WHERE tiposMovimientoCaja.id ='" & l_idtipomovimiento & "'"
	l_sql = l_sql & " and tiposMovimientoCaja.flagcompra = -1 " 
	
	l_sql = l_sql & " and tiposMovimientoCaja.empnro = " & Session("empnro")   

	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof and l_idcompraorigen = "0" then
	    texto =  "El tipo de movimiento es "& l_rs("descripcion") &".</br>"&"</br>"
		texto =  texto& "Debe seleccionar una Compra."
	end if 
	l_rs.close
end if

'Chequeo si es una venta que exista saldo a cobrar
if texto = "OK" and l_idventaorigen <> "0" then
	
	l_sql =  " 	select	isnull(  "
	l_sql = l_sql & " 				(	select SUM( detalleVentas.cantidad*detalleVentas.precio_unitario)  "
	l_sql = l_sql & " 					from detalleVentas  "
	l_sql = l_sql & " 					where detalleVentas.idVenta =  "&l_idventaorigen
	l_sql = l_sql & " 				) "
	l_sql = l_sql & " 				,0)																			as monto_venta  "
	l_sql = l_sql & " 		,isnull( "
	l_sql = l_sql & " 				(	select	SUM(cajaMovimientos.monto)  "
	l_sql = l_sql & " 					from	cajaMovimientos  "
	l_sql = l_sql & " 					where	cajaMovimientos.idventaOrigen =   "&l_idventaorigen
	if l_tipo = "M" then
		l_sql = l_sql & " 		and cajaMovimientos.id <> "&l_id
	end if
	l_sql = l_sql & " 				),0)	+ "&l_monto&"																	as monto_cobrado "	

	'response.write "l_sql"& l_sql
	rsOpen l_rs, cn, l_sql, 0

	if not l_rs.eof and (cdbl(l_rs("monto_venta")) - cdbl(l_rs("monto_cobrado") ) < 0 )then	
	    texto =  "Monto de la venta: "&l_rs("monto_venta") &"</br>"
		texto =  texto& "Monto cobrado: "&l_rs("monto_cobrado")&"</br>" &"</br>"
		texto =  texto& "No tiene saldo suficiente. Revise la venta y los cobros asociados."
	end if 
	l_rs.close
end if

'Chequeo si es una compra que exista saldo a pagar
if texto = "OK" and l_idcompraorigen <> "0" then
	
	l_sql =  " 	select	isnull(  "
	l_sql = l_sql & " 				(	select SUM( detalleCompras.cantidad*detalleCompras.precio_unitario)  "
	l_sql = l_sql & " 					from detalleCompras  "
	l_sql = l_sql & " 					where detalleCompras.idcompra =  "&l_idcompraorigen
	l_sql = l_sql & " 				) "
	l_sql = l_sql & " 				,0)																			as monto_compra  "
	l_sql = l_sql & " 		,isnull( "
	l_sql = l_sql & " 				(	select	SUM(cajaMovimientos.monto)  "
	l_sql = l_sql & " 					from	cajaMovimientos  "
	l_sql = l_sql & " 					where	cajaMovimientos.idcompraorigen =   "&l_idcompraorigen
	if l_tipo = "M" then
		l_sql = l_sql & " 		and cajaMovimientos.id <> "&l_id
	end if
	l_sql = l_sql & " 				),0)	+ "&l_monto&"																	as monto_pagado "	

	'response.write "l_sql"& l_sql
	rsOpen l_rs, cn, l_sql, 0

	if not l_rs.eof and (cdbl(l_rs("monto_compra")) - cdbl(l_rs("monto_pagado") ) < 0 )then	
	    texto =  "Monto de la compra: "&l_rs("monto_compra") &"</br>"
		texto =  texto& "Monto cobrado: "&l_rs("monto_pagado")&"</br>" &"</br>"
		texto =  texto& "No tiene saldo suficiente. Revise la compra y los pagos asociados."
	end if 
	l_rs.close
end if
%>

<% Response.write texto %>

<%
Set l_rs = Nothing
%>

