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
dim l_nomb_concept_cv
dim l_id_concept_cv
dim l_cantidad
dim l_precio

l_id_new_emp = 65

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
' ------------------------------------------------------------------------------------------------------------------
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
' ------------------------------------------------------------------------------------------------------------------
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





' ------------------------------------------------------------------------------------------------------------------
Set l_cm = Nothing

cn.CommitTrans 


Response.write "FIN "& "<br>"
%>

