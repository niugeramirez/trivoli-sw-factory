<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

dim l_id

dim l_fecha
dim l_idclientepaciente  
dim l_idrecursoreservable
dim l_detalle



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_fecha			           = request.Form("fecha")
l_detalle 		           = ConvertFromUTF8(replace(request.Form("detalle"),vbCrlf,"</br>"))'replace(request.Form("detalle"),vbCrlf,"</br>")  'request.Form("detalle")
l_idrecursoreservable 	   = request.Form("idrecursoreservable")

l_idclientepaciente = request.Form("idclientepaciente") '199624



if len(l_fecha) = 0 then
	l_fecha = "null"
else 
	l_fecha = cambiafecha(l_fecha,"YMD",true)	
end if 

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO historia_clinica_resumida  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (fecha, idclientepaciente, idrecursoreservable, detalle,empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES (" & l_fecha & "," & l_idclientepaciente & "," & l_idrecursoreservable & ",'" & l_detalle & "','"& session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE historia_clinica_resumida "
		l_sql = l_sql & " SET idclientepaciente    = " & l_idclientepaciente & ""
		l_sql = l_sql & "    ,fecha    			   = " & l_fecha & ""	
		l_sql = l_sql & "    ,idrecursoreservable    = " & l_idrecursoreservable & ""
		l_sql = l_sql & "    ,detalle    =    '" & l_detalle & "'"
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

