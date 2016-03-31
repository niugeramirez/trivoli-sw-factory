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

dim l_fecha_desde 
dim l_fecha_hasta 
dim l_idconceptoCompraVenta 
dim l_cantidadproyectada


l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")

l_fecha_desde 			= request.Form("fecha_desde")
l_fecha_hasta 			= request.Form("fecha_hasta")
l_idconceptoCompraVenta = request.Form("idconceptoCompraVenta")
l_cantidadproyectada 	= request.Form("cantidadproyectada")

if len(l_fecha_desde) = 0 then
	l_fecha_desde = "null"
else 
	l_fecha_desde = cambiafecha(l_fecha_desde,"YMD",true)	
end if 

if len(l_fecha_hasta) = 0 then
	l_fecha_hasta = "null"
else 
	l_fecha_hasta = cambiafecha(l_fecha_hasta,"YMD",true)	
end if 

'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO proyeccionVentas  "
		l_sql = l_sql & " (fecha_desde,fecha_hasta ,idconceptoCompraVenta,cantidadproyectada, empnro, created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " VALUES (" &l_fecha_desde &","& l_fecha_hasta &",'"&l_idconceptoCompraVenta&"','"&l_cantidadproyectada& "','" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE proyeccionVentas "
		l_sql = l_sql & " SET fecha_desde    = " & l_fecha_desde & ""		
		l_sql = l_sql & " , fecha_hasta    = " & l_fecha_hasta & ""
		l_sql = l_sql & " , idconceptoCompraVenta    = '" & l_idconceptoCompraVenta & "'"
		l_sql = l_sql & " , cantidadproyectada    = '" & l_cantidadproyectada & "'"
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

