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
dim l_nombre
dim l_telefono
dim l_celular  
dim l_mail  



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_nombre	               = request.Form("nombre")
l_telefono		           = request.Form("telefono")
l_celular		           = request.Form("celular")
l_mail  		           = request.Form("mail")


'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO proveedores  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (nombre,telefono,celular, mail, empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES ('" & l_nombre & "','" & l_telefono  &  "','" & l_celular  &  "','" & l_mail  &  "','" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE proveedores "
		l_sql = l_sql & " SET nombre    = '" & l_nombre & "'"
		l_sql = l_sql & "    ,telefono  = '" & l_telefono & "'"	
		l_sql = l_sql & "    ,celular   = '" & l_celular & "'"
		l_sql = l_sql & "    ,mail      = '" & l_mail & "'"
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

