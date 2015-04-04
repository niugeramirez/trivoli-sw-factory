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
dim l_apellido
dim l_nombre  
dim l_nrohistoriaclinica
dim l_dni     
dim l_domicilio
dim l_telefono
dim l_idobrasocial


l_tipo 		     = request.querystring("tipo")
l_id             = request.Form("id")
l_apellido       = request.Form("apellido")
l_nombre         = request.Form("nombre")
l_nrohistoriaclinica = request.Form("nrohistoriaclinica")
l_dni            = request.Form("dni")
l_domicilio      = request.Form("domicilio")
l_telefono       = request.Form("telefono")
'Response.write "<script>alert('Operación telefono " & l_telefono  &" Realizada.');</script>"
l_idobrasocial	 =  request.Form("osid")

'Response.write "<script>alert('Operación " & l_idobrasocial  &" Realizada.');</script>"


set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO clientespacientes "
	l_sql = l_sql & " (apellido, nombre, nrohistoriaclinica , dni,domicilio, telefono,idobrasocial ,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES ('" & l_apellido & "','" & l_nombre & "'," & l_nrohistoriaclinica & "," & l_dni & ",'" & l_domicilio & "','" & l_telefono & "'," & l_idobrasocial &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
else
	l_sql = "UPDATE clientespacientes "
	l_sql = l_sql & " SET apellido    = '" & l_apellido & "'"
	l_sql = l_sql & "    ,nombre    = '" & l_nombre & "'"
	l_sql = l_sql & "    ,nrohistoriaclinica    = " & l_nrohistoriaclinica & ""	
	l_sql = l_sql & "    ,dni    =    " & l_dni & ""
	l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
	l_sql = l_sql & "    ,telefono      = '" & l_telefono & "'"	
	l_sql = l_sql & "    ,idobrasocial  = " & l_idobrasocial 	
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

