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


l_tipo 		     = request.querystring("tipo")
l_id             = request.Form("id") ' Calendario
l_pacienteid     = request.Form("pacienteid") ' Paciente
l_apellido       = request.Form("apellido")
l_nombre         = request.Form("nombre")
l_dni            = request.Form("dni")
l_domicilio      = request.Form("domicilio")
l_tel            = request.Form("tel")
l_idobrasocial   = request.Form("osid")
l_idpractica   = request.Form("practicaid")
'l_idobrasocial      = request.Form("legape")




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
    if l_pacienteid = -1 then
	l_sql = "INSERT INTO turnos "
	l_sql = l_sql & " (idcalendario, idclientepaciente, idos, idpractica, apellido, nombre, dni, domicilio, telefono)"
	l_sql = l_sql & " VALUES (" & l_id & "," & l_pacienteid & "," & l_idobrasocial & "," & l_idpractica & ",'" & l_apellido & "','" & l_nombre & "'," & l_dni & ",'" & l_domicilio & "','" & l_tel  & "')"	
	
	else
	
	l_sql = "INSERT INTO turnos "
	l_sql = l_sql & " (idcalendario, idclientepaciente, idos, idpractica)"
	l_sql = l_sql & " VALUES (" & l_id & "," & l_pacienteid & "," & l_idobrasocial & "," & l_idpractica &  ")"
	end if
'else
'	l_sql = "UPDATE clientespacientes "
'	l_sql = l_sql & " SET apellido    = '" & l_apellido & "'"
'	l_sql = l_sql & "    ,nombre    = '" & l_nombre & "'"
'	l_sql = l_sql & "    ,dni    =    " & l_dni & ""
'	l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
'	l_sql = l_sql & " WHERE id = " & l_id
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

