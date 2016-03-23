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
dim l_idestadoInstalacion 
dim l_fechaProgramadaInstalacion


l_tipo 		               		= request.Form("tipo")
l_id                       		= request.Form("id")
l_idconceptoCompraVenta    		= request.Form("idconceptoCompraVenta")
l_cantidad				   		= request.Form("cantidad2")
l_precio_unitario		   		= request.Form("precio_unitario2")
l_observaciones            		= request.Form("observaciones")
l_idventa				   		= request.Form("idventa")

l_idestadoInstalacion 			= request.Form("estadoInstalacion")
if isnull(l_idestadoInstalacion) or len(l_idestadoInstalacion)=0 then
	l_idestadoInstalacion = 0
end if


l_fechaProgramadaInstalacion	= request.Form("fechaProgramadaInstalacion")
if len(l_fechaProgramadaInstalacion) = 0 then
	l_fechaProgramadaInstalacion = "null"
else 
	l_fechaProgramadaInstalacion = cambiafecha(l_fechaProgramadaInstalacion,"YMD",true)	
end if 

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO detalleVentas  "
		l_sql = l_sql & " (idVenta, idconceptoCompraVenta,cantidad,precio_unitario, observaciones, empnro, created_by,creation_date,last_updated_by,last_update_date, idestadoInstalacion , fechaProgramadaInstalacion)"		
		l_sql = l_sql & " VALUES (" & l_idventa & "," & l_idconceptoCompraVenta & "," & l_cantidad  &  "," & l_precio_unitario  &  ",'" & l_observaciones  &  "','" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE()," & l_idestadoInstalacion  &  "," & l_fechaProgramadaInstalacion  &  ")"
		
	else
		l_sql = "UPDATE detalleVentas "
		l_sql = l_sql & " SET idconceptoCompraVenta    = " & l_idconceptoCompraVenta & " "
		l_sql = l_sql & "    ,cantidad  = " & l_cantidad & " "	
		l_sql = l_sql & "    ,precio_unitario   = " & l_precio_unitario & " "
		l_sql = l_sql & "    ,observaciones      = '" & l_observaciones & "'"
		l_sql = l_sql & "    ,idestadoInstalacion  = " & l_idestadoInstalacion & ""			
		l_sql = l_sql & "    ,fechaProgramadaInstalacion  = " & l_fechaProgramadaInstalacion& ""			
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

