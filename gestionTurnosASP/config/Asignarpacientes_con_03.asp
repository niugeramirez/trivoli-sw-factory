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
dim l_pacienteid
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_domicilio
dim l_tel
dim l_idobrasocial
dim l_idpractica
dim l_comentario
dim l_idrecursoreservable


l_tipo 		     = request.querystring("tipo")
l_id             = request.Form("id") ' Calendario
l_pacienteid     = request.Form("pacienteid") ' Paciente
l_apellido       = request.Form("apellido")
l_nombre         = request.Form("nombre")
l_dni            = request.Form("dni")
l_domicilio      = request.Form("domicilio")
l_tel            = request.Form("tel")
l_idobrasocial   = request.Form("osid")
l_idpractica     = request.Form("practicaid")
l_comentario     = request.Form("comentario")
l_idrecursoreservable = request.Form("idrecursoreservable")




if l_pacienteid = "" then
	l_pacienteid = -1
end if
'else 
'	l_legfecing = cambiafecha(l_legfecing,"YMD",true)	
'end if 
'if len(l_legfecnac) = 0 then
'	l_legfecnac = "null"
'else 
'	l_legfecnac = cambiafecha(l_legfecnac,"YMD",true)	
'end if 

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
    'if l_pacienteid = -1 then
	'l_sql = "INSERT INTO turnos "
	'l_sql = l_sql & " (idcalendario, idclientepaciente, idpractica, apellido, nombre, dni, domicilio, telefono, comentario, idrecursoreservable)"
	'l_sql = l_sql & " VALUES (" & l_id & "," & l_pacienteid & "," & l_idobrasocial & "," & l_idpractica & ",'" & l_apellido & "','" & l_nombre & "'," & l_dni & ",'" & l_domicilio & "','" & l_tel & "','" & l_comentario  & "'," & l_idrecursoreservable & ")"	
	
	'else
	
	l_sql = "INSERT INTO turnos "
	l_sql = l_sql & " (idcalendario, idclientepaciente, idpractica, comentario , idrecursoreservable)"
	l_sql = l_sql & " VALUES (" & l_id & "," & l_pacienteid & "," & l_idpractica & ",'" & l_comentario &  "'," & l_idrecursoreservable & ")"
	'end if
else
	l_sql = "UPDATE turnos "
	l_sql = l_sql & " SET idpractica    = " & l_idpractica
	l_sql = l_sql & "    ,comentario    = '" & l_comentario & "'"
	l_sql = l_sql & "    ,idrecursoreservable  = " & l_idrecursoreservable
	l_sql = l_sql & " WHERE id = " & l_id
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

