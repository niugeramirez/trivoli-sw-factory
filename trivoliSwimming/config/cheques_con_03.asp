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
dim l_numero
dim l_fecha_emision
dim l_fecha_vencimiento  
dim l_idbanco  
dim l_importe
dim l_flag_propio
dim l_emisor



l_tipo 		               = request.Form("tipo")
l_id                       = request.Form("id")
l_numero                   = request.Form("numero")
l_fecha_emision	           = request.Form("fecha_emision")
l_fecha_vencimiento        = request.Form("fecha_vencimiento")
l_idbanco  		           = request.Form("idbanco")
l_importe		           = request.Form("importe2")
l_flag_propio			   = request.Form("flag_propio")
l_emisor			       = request.Form("emisor")
'l_cantsobreturnos          = 0 ' request.Form("cantsobreturnos") se elimino esta campo


if len(l_fecha_emision) = 0 then
	l_fecha_emision = "null"
else 
	l_fecha_emision = cambiafecha(l_fecha_emision,"YMD",true)	
end if 

if len(l_fecha_vencimiento) = 0 then
	l_fecha_vencimiento = "null"
else 
	l_fecha_vencimiento = cambiafecha(l_fecha_vencimiento,"YMD",true)	
end if 



'response.write "l_tipo"&l_tipo & "<br>"
'response.write "l_id"&l_id & "<br>"

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO cheques  "
		' Multiempresa
		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " (descripcion, idtemplatereserva, cantturnossimult, cantsobreturnos,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " (numero, fecha_emision, fecha_vencimiento, id_banco , importe, flag_propio, emisor,  empnro, created_by,creation_date,last_updated_by,last_update_date)"

		' Se elimina est linea y se reemplaza por el codigo de abajo
		'l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_idtemplatereserva & "," & l_cantturnossimult & "," & l_cantsobreturnos  &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		l_sql = l_sql & " VALUES ('" & l_numero & "'," & l_fecha_emision & "," & l_fecha_vencimiento & "," & l_idbanco & "," & l_importe & "," & l_flag_propio & ",'" & l_emisor & "','" & session("empnro") &"','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE cheques "
		l_sql = l_sql & " SET numero    = '" & l_numero & "'"
		l_sql = l_sql & "    ,fecha_emision  = " & l_fecha_emision & ""	
		l_sql = l_sql & "    ,fecha_vencimiento  = " & l_fecha_vencimiento & ""	
		l_sql = l_sql & "    ,id_banco   = " & l_idbanco & ""
		l_sql = l_sql & "    ,importe      = " & l_importe & ""
		l_sql = l_sql & "    ,flag_propio = " & l_flag_propio & ""
		l_sql = l_sql & "    ,emisor  =  '" & l_emisor & "'"		
		'l_sql = l_sql & "    ,cantturnossimult    = " & l_cantturnossimult & ""
		'l_sql = l_sql & "    ,cantsobreturnos    =    " & l_cantsobreturnos & ""
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

