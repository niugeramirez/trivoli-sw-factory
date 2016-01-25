<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

dim l_id
dim l_idconceptoCompraVenta
dim l_cantidad
dim l_precio_unitario
dim l_observaciones

Dim l_idventa



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")

l_idconceptoCompraVenta    = request.Form("idconceptoCompraVenta")
l_cantidad				   = request.Form("cantidad")
l_precio_unitario		   = request.Form("precio_unitario")
l_observaciones            = request.Form("observaciones")

l_idventa				   = request.Form("idventa")


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO detalleVentas  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (idVenta, idconceptoCompraVenta,cantidad,precio_unitario, observaciones, empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES (" & l_idventa & "," & l_idconceptoCompraVenta & "," & l_cantidad  &  "," & l_precio_unitario  &  ",'" & l_observaciones  &  "','" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE detalleVentas "
		l_sql = l_sql & " SET idconceptoCompraVenta    = " & l_idconceptoCompraVenta & " "
		l_sql = l_sql & "    ,cantidad  = " & l_cantidad & " "	
		l_sql = l_sql & "    ,precio_unitario   = " & l_precio_unitario & " "
		l_sql = l_sql & "    ,observaciones      = '" & l_observaciones & "'"

		l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
		l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
		l_sql = l_sql & " WHERE id = " & l_id
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "OK"
%>

