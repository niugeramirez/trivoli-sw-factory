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
dim l_idmedicoderivador
dim l_iduser
Dim l_agenda


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
l_idmedicoderivador = request.Form("idmedicoderivador")
l_iduser         = request.Form("iduser")
l_agenda         = request.Form("agenda")


if l_pacienteid = "" then
	l_pacienteid = -1
end if
if l_idmedicoderivador = "" then
	l_idmedicoderivador = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	
	l_sql = "INSERT INTO turnos "
	l_sql = l_sql & " (idcalendario, idclientepaciente, idpractica, comentario , idmedicoderivador, iduseringresoturno, empnro, created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES (" & l_id & "," & l_pacienteid & "," & l_idpractica & ",'" & l_comentario &  "'," & l_idmedicoderivador & ",'" & l_iduser & "','" & session("empnro") & "','" & session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"

else
	l_sql = "UPDATE turnos "
	l_sql = l_sql & " SET idpractica    = " & l_idpractica
	l_sql = l_sql & "    ,comentario    = '" & l_comentario & "'"
	l_sql = l_sql & "    ,idmedicoderivador  = " & l_idmedicoderivador
	l_sql = l_sql & "    ,iduseringresoturno  = '" & l_iduser & "'"
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 
	l_sql = l_sql & " WHERE id = " & l_id
end if
'response.write l_sql & "<br>"
Response.write "<script>alert('Operación "&l_sql&" Realizada.');</script>"

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing
if l_agenda = "S" then
	Response.write "<script>alert('Operación Realizada .');window.parent.opener.parent.ifrm2.location.reload();window.parent.close();</script>"
else
	Response.write "<script>alert('Operación Realizada .');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
end if
%>

