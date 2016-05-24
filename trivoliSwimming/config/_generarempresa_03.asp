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

Dim l_empresa
dim l_id_new_emp
dim l_id_new_prov

l_empresa = "Prueba"
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

'EMPRESA
' ------------------------------------------------------------------------------------------------------------------
Response.write "insert empresa "& "<br>"

l_sql = "insert into empresa values ('"&l_empresa&"')"
 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_id_new_emp = codigogenerado("empresa")
response.write "id empresa generado "&l_id_new_emp & "<br>"

'BANCOS
' ------------------------------------------------------------------------------------------------------------------

l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Banco Rio'			, 'sa', GETDATE(), 'sa', GETDATE(), "&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Santander Rio'		, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Pcia Bs As'		, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Banco Santa Cruz'	, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ICBC'				, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Frances'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Patagonia'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'HSBC'				, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Nacion'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Credicoop'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Banco de Chubut'	, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Banco de La Pampa'	, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Banco de Neuquen'	, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Galicia'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Citibank'			, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [bancos] ([nombre_banco], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'Macro'				, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

response.write "<br>"
response.write "<br>"
response.write "<br>"
'PROVINCIAS Y CIUDADES
' ------------------------------------------------------------------------------------------------------------------

l_sql = "select * from provincias where provincias.empnro = 1"
rsOpenCursor l_rs, cn, l_sql, 1, 1
do until l_rs.eof
	response.write l_rs("provincia") & "<br>"
	l_sql = "INSERT [provincias] ([provincia], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('"&l_rs("provincia")&"', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
	response.write l_sql & "<br>"	
	ejecutar_sql(l_sql)
	l_id_new_prov = codigogenerado("empresa")
	response.write "id provincia generado "&l_id_new_prov & "<br>"
	
	l_sql = "select * from ciudades where ciudades.empnro = 1 and ciudades.idProvincia = "&l_rs("id")
	response.write l_sql & "<br>"
	rsOpenCursor l_rs2, cn, l_sql, 1, 1
	'if not l_rs2.eof then
	do until l_rs2.eof
		response.write l_rs2("ciudad") & "<br>"
		l_sql = "INSERT [ciudades] ( [ciudad], [codigo_postal], [idProvincia], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('"&l_rs2("ciudad")&"', '"&l_rs2("codigo_postal")&"',"& l_id_new_prov&", 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")" 
		response.write l_sql & "<br>"	
		ejecutar_sql(l_sql)		
		l_rs2.MoveNext
	loop
	'end if
	l_rs2.close

	response.write "<br>"
	l_rs.MoveNext
loop
l_rs.close

'CONCEPTOS COMPRA VENTA
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'KOMODO' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ITACARE' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'GENOVA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'HIDROS' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'BOQUILLA DE CALEFACCION', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'LUZ', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'BOMBA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'POZO E INSTALACION', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PIEDRA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CONTENEDOR' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CARRETILLADO' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CAÑERÍAS' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ELECTRICIDAD', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CODOS 90' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CUPLAS'  , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'VARIOS', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'GRUA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CERCO BABYSAFE' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PUERTA MAGNETICA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'INSTALACION DE CERCO BABYSAFE' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'COMISIONES' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'IVA' , 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'OTROS IMPUESTOS', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'BRINDICE', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'AMARALINA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ARMACAO', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ATENAS', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'ATLANTIDA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'BALDOSAS', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'BELIZE', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'CALARI', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'FAROL DA BARRA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'FIJI', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'FLORENZA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'HALEIWA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'HANALEI', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'HAPUNA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'MAZARA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PESCARA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PUPUKEA', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PARTE B', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

l_sql = "INSERT [conceptosCompraVenta] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ( 'PUBLICIDAD', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
'ESTADO INSTALACION
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [estadoInstalacion] ( [codigo], [descripcionEstadoInsta], [orden], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('P', 'Programada',1, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [estadoInstalacion] ( [codigo], [descripcionEstadoInsta], [orden], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('E', 'En Curso',2, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [estadoInstalacion] ( [codigo], [descripcionEstadoInsta], [orden], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('F', 'Finalizada',3, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)	


'MEDIOS DE PAGO
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Efectivo', NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Cheque', -1, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Pagare', NULL,'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Tarjeta Credito', NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Tarjeta Debito', NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Mercado Pago', NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [mediosdepago] ( [titulo], [flag_cheque], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Transferencia', NULL,'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)


'RESPONSABLES CAJA
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [responsablesCaja] ([iniciales], [nombre], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('F', 'Franquiciado','sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'TIPOS DE MOVIMIENTO CAJA
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Venta', 'Venta', NULL, -1, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Proveedores', 'Compra', -1, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Inversiones', 'Compra', NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Gastos', 'Compra', NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Empleados', 'Otros', NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [tiposMovimientoCaja] ( [descripcion], [origen], [flagCompra], [flagVenta], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Prestamos', 'Otros', NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'UNIDADES DE NEGOCIO
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [unidadesNegocio] ([descripcion], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('Igui', 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)

'PROVEEDORES
' ------------------------------------------------------------------------------------------------------------------
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('iGUi ALVEAR PISCINAS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('iGUi PROGEU BOMBAS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('PISOS ATERMICOS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('PISOS DE MADERA', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('PISOS Y REVESTIENTOS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('MADERERAS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('LONERIA', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('ALBAÑILES', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('AYUDANTES DE ALBAÑIL', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('CERCO REMOVIBLE', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('INSUMOS DE PISCINAS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('FLETES', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('FRANQUICIAS iGUi', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('GASTOS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('ALQUILER', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('GESTORIA', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('LUZ', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('GAS', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('TELEFONIA', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('INTERNET', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
response.write l_sql & "<br>"
ejecutar_sql(l_sql)
l_sql = "INSERT [proveedores] ([nombre], [telefono], [celular], [mail], [created_by], [creation_date], [last_updated_by], [last_update_date], [empnro]) VALUES ('MUNICIPALIDAD', NULL, NULL, NULL, 'sa', GETDATE(), 'sa', GETDATE(),"&l_id_new_emp&")"
' ------------------------------------------------------------------------------------------------------------------
Set l_cm = Nothing

cn.CommitTrans 


Response.write "FIN "& "<br>"
%>

