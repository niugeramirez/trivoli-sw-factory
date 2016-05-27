<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/adovbs.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/sqls.inc"-->
<% 


on error goto 0


Dim l_cm
Dim l_sql
Dim l_rs

Dim l_sql2
Dim l_rs2


dim l_id_new_emp
dim l_id_proveedor
dim l_id_compras
dim l_id_compras_reentrega_cheque
dim l_nomb_concept_cv
dim l_id_concept_cv
dim l_cantidad
dim l_precio
dim l_id_cliente
dim l_id_ventas
dim l_idtipoMovimiento_ventas
dim l_idtipoMovimiento_prov
dim l_idunidadNegocio
dim l_idmedioPago_cheque
dim l_idmedioPago_eft
dim l_idresponsable
dim l_id_banco
dim l_id_banco_02
dim l_id_cheque
dim l_id_cheque_reentrega
dim l_idestadoInstalacion_prog

l_id_new_emp = 66

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")


' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado(tabla)
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id",tabla)
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()

' ------------------------------------------------------------------------------------------------------------------
' ejecutar_sql() :
' ------------------------------------------------------------------------------------------------------------------
sub ejecutar_sql(sql)

	dim l_cm
	set l_cm = Server.CreateObject("ADODB.Command")
	
	l_cm.activeconnection = Cn
	l_cm.CommandText = sql
	cmExecute l_cm, sql, 0	
end sub 

'Al operar sobre varias tablas debo iniciar una transacción
cn.BeginTrans
 
Response.write "INICIO"& "<br>"

'COMPRA PISCINAS
'*******************************************************************************************************************
l_sql = "select * from proveedores where UPPER(proveedores.nombre) = 'IGUI ALVEAR PISCINAS' and proveedores.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_proveedor = l_rs("id")
	response.write "Proveedor id "&l_id_proveedor & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [compras] ( [fecha], [idproveedor], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES (cast(getDate()-30 As Date),"& l_id_proveedor &", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_id_compras = codigogenerado("compras")
response.write "id compras generado "&l_id_compras & "<br>"
l_id_compras_reentrega_cheque = l_id_compras



'DETALLE COMPRA BRINDICE
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "BRINDICE"
l_cantidad = 1
l_precio = 41268

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE COMPRA CALARI
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "CALARI"
l_cantidad = 5
l_precio = 34388

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)


'DETALLE COMPRA FAROL DA BARRA
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "FAROL DA BARRA"
l_cantidad = 1
l_precio = 27514

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE COMPRA FLORENZA
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "FLORENZA"
l_cantidad = 5
l_precio = 26135

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
		

'DETALLE COMPRA CALARI
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "MAZARA"
l_cantidad = 1
l_precio = 20084

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE COMPRA BOMBA
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "BOMBA"
l_cantidad = 13
l_precio = 15030

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE COMPRA VARIOS PUBLICIDAD
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "PUBLICIDAD"
l_cantidad = 1
l_precio = 64045

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE COMPRA PARTE B
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "PARTE B"
l_cantidad = 1
l_precio = 25000

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
	
'COMPRA BOMBAS CAMBIO
'*******************************************************************************************************************
l_sql = "select * from proveedores where UPPER(proveedores.nombre) = 'IGUI PROGEU BOMBAS' and proveedores.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_proveedor = l_rs("id")
	response.write "Proveedor id "&l_id_proveedor & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [compras] ( [fecha], [idproveedor], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES (cast(getDate()-30 As Date),"& l_id_proveedor &", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_id_compras = codigogenerado("compras")
response.write "id compras generado "&l_id_compras & "<br>"

'DETALLE COMPRA BOMBA CAMBIO
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "BOMBA"
l_cantidad = 13
l_precio = 5030

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleCompras] ( [idcompra], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_compras&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'COMPRRA DE PRUEBA CAMBIO DE  BOMBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'VENTA
'*******************************************************************************************************************
l_sql = "select * from [clientes] where clientes.nombre = 'CLIENTE DE PRUEBA: Eugenio Ramirez' and clientes.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_cliente = l_rs("id")
	response.write "Cliente id "&l_id_cliente & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [ventas] ( [fecha], [idcliente], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES (cast(getDate()-15 As Date),"& l_id_cliente &", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_id_ventas = codigogenerado("ventas")
response.write "id ventas generado "&l_id_ventas & "<br>"
'INICIALIZACIONES DETALLE VENTA 
' ------------------------------------------------------------------------------------------------------------------
l_sql = "select * from estadoInstalacion where codigo = 'P'and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idestadoInstalacion_prog = l_rs("id")
	response.write "Estado instalacion programada id "&l_idestadoInstalacion_prog & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

'DETALLE VENTA BOQUILLA DE CALEFACCION
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "BOQUILLA DE CALEFACCION"
l_cantidad = 1
l_precio = 800

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO BOQUILLA DE CALEFACCION
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"400," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE VENTA FLORENZA
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "FLORENZA"
l_cantidad = 1
l_precio = 67000

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [idestadoInstalacion],[fechaProgramadaInstalacion],[created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'," 'observaciones
l_sql = l_sql&l_idestadoInstalacion_prog&"," '[idestadoInstalacion]
l_sql = l_sql&"cast(getDate()+20 As Date)" ',[fechaProgramadaInstalacion]
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO VENTA FLORENZA
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"26135," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE VENTA HIDROS
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "HIDROS"
l_cantidad = 1
l_precio = 1200

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO HIDROS
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"100," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE VENTA LUZ
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "LUZ"
l_cantidad = 1
l_precio = 3000

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO LUZ
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"2000," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE VENTA PIEDRA
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "PIEDRA"
l_cantidad = 1
l_precio = 7000

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO PIEDRA
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"4500," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'DETALLE VENTA POZO E INSTALACION
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "POZO E INSTALACION"
l_cantidad = 1
l_precio = 12000

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT [detalleVentas] ( [idventa], [idconceptoCompraVenta], [cantidad], [precio_unitario], [observaciones], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql& l_id_ventas&"," '[idcompra]
l_sql = l_sql&l_id_concept_cv&"," '[idconceptoCompraVenta]
l_sql = l_sql&l_cantidad&"," '[cantidad]
l_sql = l_sql&l_precio&"," 'precio_unitario
l_sql = l_sql&"'VENTA DE PRUEBA'" 'observaciones
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO POZO E INSTALACION
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&l_cantidad&"," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"9000," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO VARIOS 
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "VARIOS"

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&"1," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"800," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA: Viajes / Gasoil / Viaticos'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'COSTO BOMBA 
' ------------------------------------------------------------------------------------------------------------------
l_nomb_concept_cv = "BOMBA"

l_sql = "select * from conceptosCompraVenta where conceptosCompraVenta.descripcion = '"&l_nomb_concept_cv&"' and  conceptosCompraVenta.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_concept_cv = l_rs("id")
	response.write l_nomb_concept_cv&" id "&l_id_concept_cv & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "INSERT INTO [costosVentas]([idVenta],[idconceptoCompraVenta],[cantidad],[precio_unitario],[observaciones],[created_by],[creation_date],[last_updated_by],[last_update_date],[empnro])VALUES("
l_sql = l_sql&l_id_ventas&"," '<idVenta, int,>
l_sql = l_sql&l_id_concept_cv&"," ',<idconceptoCompraVenta, int,>
l_sql = l_sql&"1," ',<cantidad, decimal(19,4),>
l_sql = l_sql&"15030," ',<precio_unitario, decimal(19,4),>
l_sql = l_sql&"'COSTO DE PRUEBA: Viajes / Gasoil / Viaticos'" ',<observaciones, varchar(100),>
l_sql = l_sql&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'PAGOS VENTA
'*******************************************************************************************************************

'INICIALIZACIONES 
' ------------------------------------------------------------------------------------------------------------------
l_sql = "select * from [tiposMovimientoCaja] where flagVenta = -1 and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idtipoMovimiento_ventas = l_rs("id")
	response.write "Tipo movimiento venta id "&l_idtipoMovimiento_ventas & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from [tiposMovimientoCaja] where flagCompra = -1 and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idtipoMovimiento_prov = l_rs("id")
	response.write "Tipo movimiento proveedores id "&l_idtipoMovimiento_prov & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from unidadesNegocio where unidadesNegocio.empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idunidadNegocio = l_rs("id")
	response.write "Unidade de negocio id "&l_idunidadNegocio & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from mediosdepago where titulo = 'Cheque' and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idmedioPago_cheque = l_rs("id")
	response.write "Medio de pago cheque id "&l_idmedioPago_cheque & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from mediosdepago where titulo = 'Efectivo' and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idmedioPago_eft = l_rs("id")
	response.write "Medio de pago efectivo id "&l_idmedioPago_eft & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from responsablesCaja where empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_idresponsable = l_rs("id")
	response.write "Responsable caja id "&l_idresponsable & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from bancos  where nombre_banco = 'HSBC' and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_banco = l_rs("id")
	response.write "Banco id "&l_id_banco & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

l_sql = "select * from bancos  where nombre_banco = 'Santander Rio' and empnro = "&l_id_new_emp
rsOpenCursor l_rs, cn, l_sql, 1, 1

do until l_rs.eof
	l_id_banco_02 = l_rs("id")
	response.write "Banco 2 id "&l_id_banco_02 & "<br>"
	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close
'PRIMER PAGO VENTA (EFECTIVO)
' ------------------------------------------------------------------------------------------------------------------


l_sql = "INSERT [cajaMovimientos] ([fecha], [tipo], [idtipoMovimiento], [detalle], [idunidadNegocio], [idmedioPago], [idcheque], [monto], [idresponsable], [idcompraOrigen], [idventaOrigen], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro])VALUES ( "
l_sql = l_sql& "cast(getDate()-15  as Date)," '[fecha]
l_sql = l_sql&"'E'," ' [tipo]
l_sql = l_sql&l_idtipoMovimiento_ventas&"," ' [idtipoMovimiento]
l_sql = l_sql&"'COBRO DE PRUEBA'," ' [detalle]
l_sql = l_sql&l_idunidadNegocio&"," ' [idunidadNegocio]
l_sql = l_sql&l_idmedioPago_eft&"," ' [idmedioPago]
l_sql = l_sql&"0," ' [idcheque]
l_sql = l_sql&"7000," ' [monto]
l_sql = l_sql&l_idresponsable&"," ' [idresponsable]
l_sql = l_sql&"0," ' [idcompraOrigen]
l_sql = l_sql&l_id_ventas&"," ' [idventaOrigen]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"

response.write l_sql & "<br>"
ejecutar_sql(l_sql)



'SEGUNDA PAGO VENTA (CHEQUE)
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [cheques] ( [numero], [fecha_emision], [fecha_vencimiento], [id_banco], [importe], [flag_emitidopor_cliente], [emisor], [flag_propio], [validacion_bcra], [flag_cobrado_pagado], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql&"'PRUEBA84854780'," ' [numero]
l_sql = l_sql&"cast(getDate()-15  as Date)," ' [fecha_emision]
l_sql = l_sql&"cast(getDate()+45  as Date)," ' [fecha_vencimiento]
l_sql = l_sql&l_id_banco&"," ' [id_banco]
l_sql = l_sql&"65130," ' [importe]
l_sql = l_sql&"-1," ' [flag_emitidopor_cliente]
l_sql = l_sql&"''," ' [emisor]
l_sql = l_sql&"null," ' [flag_propio]
l_sql = l_sql&"'VALIDADO'," ' [validacion_bcra]
l_sql = l_sql&"null," '  [flag_cobrado_pagado]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_id_cheque = codigogenerado("cheques")
response.write "id cheques generado "&l_id_cheque & "<br>"
l_id_cheque_reentrega = l_id_cheque

l_sql = "INSERT [cajaMovimientos] ([fecha], [tipo], [idtipoMovimiento], [detalle], [idunidadNegocio], [idmedioPago], [idcheque], [monto], [idresponsable], [idcompraOrigen], [idventaOrigen], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro])VALUES ( "
l_sql = l_sql& "cast(getDate()-15  as Date)," '[fecha]
l_sql = l_sql&"'E'," ' [tipo]
l_sql = l_sql&l_idtipoMovimiento_ventas&"," ' [idtipoMovimiento]
l_sql = l_sql&"'COBRO DE PRUEBA'," ' [detalle]
l_sql = l_sql&l_idunidadNegocio&"," ' [idunidadNegocio]
l_sql = l_sql&l_idmedioPago_cheque &"," ' [idmedioPago]
l_sql = l_sql&l_id_cheque_reentrega&"," ' [idcheque]
l_sql = l_sql&"65130," ' [monto]
l_sql = l_sql&l_idresponsable&"," ' [idresponsable]
l_sql = l_sql&"0," ' [idcompraOrigen]
l_sql = l_sql&l_id_ventas&"," ' [idventaOrigen]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"

response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'PAGOS COMPRA
'*******************************************************************************************************************

'PRIMER PAGO CHEQUE DE CLIENTE 
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [cajaMovimientos] ([fecha], [tipo], [idtipoMovimiento], [detalle], [idunidadNegocio], [idmedioPago], [idcheque], [monto], [idresponsable], [idcompraOrigen], [idventaOrigen], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro])VALUES ( "
l_sql = l_sql& "cast(getDate()-15  as Date)," '[fecha]
l_sql = l_sql&"'S'," ' [tipo]
l_sql = l_sql&l_idtipoMovimiento_prov &"," ' [idtipoMovimiento]
l_sql = l_sql&"'PAGO DE PRUEBA'," ' [detalle]
l_sql = l_sql&l_idunidadNegocio&"," ' [idunidadNegocio]
l_sql = l_sql&l_idmedioPago_cheque &"," ' [idmedioPago]
l_sql = l_sql&l_id_cheque_reentrega&"," ' [idcheque]
l_sql = l_sql&"65130," ' [monto]
l_sql = l_sql&l_idresponsable&"," ' [idresponsable]
l_sql = l_sql&l_id_compras_reentrega_cheque& "," ' [idcompraOrigen]
l_sql = l_sql&"0," ' [idventaOrigen]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"

response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'SEGUNDO PAGO EFECTIVO
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [cajaMovimientos] ([fecha], [tipo], [idtipoMovimiento], [detalle], [idunidadNegocio], [idmedioPago], [idcheque], [monto], [idresponsable], [idcompraOrigen], [idventaOrigen], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro])VALUES ( "
l_sql = l_sql& "cast(getDate()-14  as Date)," '[fecha]
l_sql = l_sql&"'S'," ' [tipo]
l_sql = l_sql&l_idtipoMovimiento_prov &"," ' [idtipoMovimiento]
l_sql = l_sql&"'PAGO DE PRUEBA'," ' [detalle]
l_sql = l_sql&l_idunidadNegocio&"," ' [idunidadNegocio]
l_sql = l_sql&l_idmedioPago_eft &"," ' [idmedioPago]
l_sql = l_sql&"0," ' [idcheque]
l_sql = l_sql&"144876," ' [monto]
l_sql = l_sql&l_idresponsable&"," ' [idresponsable]
l_sql = l_sql&l_id_compras_reentrega_cheque& "," ' [idcompraOrigen]
l_sql = l_sql&"0," ' [idventaOrigen]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"

response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'TERCER PAGO CHEQUE PROPIO
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [cheques] ( [numero], [fecha_emision], [fecha_vencimiento], [id_banco], [importe], [flag_emitidopor_cliente], [emisor], [flag_propio], [validacion_bcra], [flag_cobrado_pagado], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ("
l_sql = l_sql&"'PRUEBA00000032'," ' [numero]
l_sql = l_sql&"cast(getDate()-13  as Date)," ' [fecha_emision]
l_sql = l_sql&"cast(getDate()+77  as Date)," ' [fecha_vencimiento]
l_sql = l_sql&l_id_banco_02&"," ' [id_banco]
l_sql = l_sql&"400000," ' [importe]
l_sql = l_sql&"null," ' [flag_emitidopor_cliente]
l_sql = l_sql&"''," ' [emisor]
l_sql = l_sql&"-1," ' [flag_propio]
l_sql = l_sql&"'VALIDADO'," ' [validacion_bcra]
l_sql = l_sql&"null," '  [flag_cobrado_pagado]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_id_cheque = codigogenerado("cheques")
response.write "id cheques generado "&l_id_cheque & "<br>"

l_sql = "INSERT [cajaMovimientos] ([fecha], [tipo], [idtipoMovimiento], [detalle], [idunidadNegocio], [idmedioPago], [idcheque], [monto], [idresponsable], [idcompraOrigen], [idventaOrigen], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro])VALUES ( "
l_sql = l_sql& "cast(getDate()-13  as Date)," '[fecha]
l_sql = l_sql&"'S'," ' [tipo]
l_sql = l_sql&l_idtipoMovimiento_prov &"," ' [idtipoMovimiento]
l_sql = l_sql&"'PAGO DE PRUEBA'," ' [detalle]
l_sql = l_sql&l_idunidadNegocio&"," ' [idunidadNegocio]
l_sql = l_sql&l_idmedioPago_cheque&"," ' [idmedioPago]
l_sql = l_sql&l_id_cheque&"," ' [idcheque]
l_sql = l_sql&"400000," ' [monto]
l_sql = l_sql&l_idresponsable&"," ' [idresponsable]
l_sql = l_sql&l_id_compras_reentrega_cheque& "," ' [idcompraOrigen]
l_sql = l_sql&"0," ' [idventaOrigen]
l_sql = l_sql&" 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"

response.write l_sql & "<br>"
ejecutar_sql(l_sql)
' ------------------------------------------------------------------------------------------------------------------
Set l_cm = Nothing

cn.CommitTrans 


Response.write "FIN "& "<br>"
%>

