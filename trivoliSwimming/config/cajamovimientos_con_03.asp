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
dim l_fecha
dim l_tipoes
dim l_idtipomovimiento  
dim l_detalle
dim l_idunidadnegocio 
dim l_idmediopago
dim l_idcheque
dim l_monto
dim l_idresponsable
dim l_idcompraorigen
dim l_idventaorigen



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_fecha                    = request.Form("fecha")
l_tipoes		           = request.Form("tipoes")
l_idtipomovimiento        = request.Form("idtipomovimiento")
l_detalle  		           = request.Form("detalle")
l_idunidadnegocio          = request.Form("idunidadnegocio")
l_idmediopago			   = request.Form("idmediopago")
l_idcheque			       = request.Form("idcheque")
l_monto				       = request.Form("monto2")
l_idresponsable 		   = request.Form("idresponsable")
l_idcompraorigen		   = request.Form("idcompraorigen")
l_idventaorigen		       = request.Form("idventaorigen") 

if len(l_fecha) = 0 then
	l_fecha = "null"
else 
	l_fecha = cambiafecha(l_fecha,"YMD",true)	
end if 

if len(l_idcheque)  = 0  then
	l_idcheque = 0
end if

if len(l_idventaorigen) = 0 then
	l_idventaorigen = "0"
end if

if len(l_idcompraorigen) = 0 then
	l_idcompraorigen = "0"
end if

if len(l_idunidadnegocio) = 0 then
	l_idunidadnegocio = "0"
end if


'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO cajaMovimientos  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (fecha, tipo, idtipomovimiento, detalle, idunidadnegocio, idmediopago, idcheque, monto, idresponsable, idcompraorigen, idventaorigen, empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES (" & l_fecha & ",'" & l_tipoes & "'," & l_idtipomovimiento & ",'" & l_detalle & "'," & l_idunidadnegocio & "," & l_idmediopago & "," & l_idcheque & "," & l_monto & "," & l_idresponsable & "," & l_idcompraorigen & "," & l_idventaorigen & ",'" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE cajaMovimientos "
		l_sql = l_sql & " SET fecha    = " & l_fecha & ""
		l_sql = l_sql & "    ,tipo  = '" & l_tipoes & "'"	
		l_sql = l_sql & "    ,idtipomovimiento  = " & l_idtipomovimiento & ""	
		l_sql = l_sql & "    ,detalle   = '" & l_detalle & "'"
		l_sql = l_sql & "    ,idunidadnegocio      = " & l_idunidadnegocio & ""
		l_sql = l_sql & "    ,idmediopago = " & l_idmediopago & ""
		l_sql = l_sql & "    ,idcheque  =  " & l_idcheque & ""		
		l_sql = l_sql & "    ,monto    = " & l_monto & ""
		l_sql = l_sql & "    ,idresponsable    =    " & l_idresponsable & ""
		l_sql = l_sql & "    ,idcompraorigen    =    " & l_idcompraorigen & ""
		l_sql = l_sql & "    ,idventaorigen    =    " & l_idventaorigen & ""
		
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

