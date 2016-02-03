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
dim l_descripcion
dim l_idtemplatereserva
dim l_cantturnossimult  
dim l_cantsobreturnos     



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_descripcion              = request.Form("descripcion")
l_idtemplatereserva        = request.Form("idtemplatereserva")
l_cantturnossimult         = request.Form("cantturnossimult")
l_cantsobreturnos          = 0 ' request.Form("cantsobreturnos") se elimino esta campo

'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO recursosreservables  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos & ",'"& session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE recursosreservables "
		l_sql = l_sql & " SET descripcion    = '" & l_descripcion & "'"
		l_sql = l_sql & "    ,idtemplatereserva    = " & l_idtemplatereserva & ""	
		l_sql = l_sql & "    ,cantturnossimult    = " & l_cantturnossimult & ""
		l_sql = l_sql & "    ,cantsobreturnos    =    " & l_cantsobreturnos & ""
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

