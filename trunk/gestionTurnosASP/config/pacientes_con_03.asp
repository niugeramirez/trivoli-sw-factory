<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql

dim l_id
dim l_apellido
dim l_nombre  
dim l_dni     
dim l_domicilio
dim l_idobrasocial


l_tipo 		     = request.querystring("tipo")
l_id             = request.Form("id")
l_apellido       = request.Form("apellido")
l_nombre         = request.Form("nombre")
l_dni            = request.Form("dni")
l_domicilio      = request.Form("domicilio")
'l_idobrasocial      = request.Form("legape")




'if len(l_legfecing) = 0 then
'	l_legfecing = "null"
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
	l_sql = "INSERT INTO clientespacientes "
	l_sql = l_sql & " (apellido, nombre, dni,domicilio)"
	l_sql = l_sql & " VALUES ('" & l_apellido & "','" & l_nombre & "'," & l_dni & ",'" & l_domicilio & "')"
else
	l_sql = "UPDATE clientespacientes "
	l_sql = l_sql & " SET apellido    = '" & l_apellido & "'"
	l_sql = l_sql & "    ,nombre    = '" & l_nombre & "'"
	l_sql = l_sql & "    ,dni    =    " & l_dni & ""
	l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
	l_sql = l_sql & " WHERE id = " & l_id
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

