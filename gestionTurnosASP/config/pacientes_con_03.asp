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

dim l_fecha_ingreso
Dim l_fechanacimiento
dim l_nro_obra_social
Dim l_sexo
Dim l_ciudad

dim l_observaciones


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

if len(l_idobrasocial) = 0 then
	l_idobrasocial = 0
end if 

l_fecha_ingreso       = request.Form("fecha_ingreso")
l_fechanacimiento       = request.Form("fechanacimiento")

if len(l_fecha_ingreso) = 0 then
	l_fecha_ingreso = "null"
else 
	l_fecha_ingreso = cambiafecha(l_fecha_ingreso,"YMD",true)	
end if 

if len(l_fechanacimiento) = 0 then
	l_fechanacimiento = "null"
else 
	l_fechanacimiento = cambiafecha(l_fechanacimiento,"YMD",true)	
end if 



l_nro_obra_social = request.Form("nro_obra_social")
l_sexo = request.Form("sexo")
l_ciudad = request.Form("ciudad")

if len(l_ciudad) = 0 then
	l_ciudad = 0
end if 

l_observaciones = request.Form("observaciones")


set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO clientespacientes "
	l_sql = l_sql & " (apellido, nombre, nrohistoriaclinica , dni,domicilio, telefono,idobrasocial, fecha_ingreso, fechanacimiento, nro_obra_social, sexo, idciudad , observaciones ,empnro,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES ('" & l_apellido & "','" & l_nombre & "','" & l_nrohistoriaclinica & "'," & l_dni & ",'" & l_domicilio & "','" & l_telefono & "'," & l_idobrasocial & "," & l_fecha_ingreso & "," & l_fechanacimiento & ",'" & l_nro_obra_social & "','" & l_sexo & "'," & l_ciudad & ",'" & l_observaciones & "','" & session("empnro") & "','" & session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
else
	l_sql = "UPDATE clientespacientes "
	l_sql = l_sql & " SET apellido    = '" & l_apellido & "'"
	l_sql = l_sql & "    ,nombre    = '" & l_nombre & "'"
	l_sql = l_sql & "    ,nrohistoriaclinica    = '" & l_nrohistoriaclinica & "'"	
	l_sql = l_sql & "    ,dni    =    " & l_dni & ""
	l_sql = l_sql & "    ,domicilio     = '" & l_domicilio & "'"
	l_sql = l_sql & "    ,telefono      = '" & l_telefono & "'"	
	l_sql = l_sql & "    ,idobrasocial  = " & l_idobrasocial 	
	l_sql = l_sql & "    ,sexo          = '" & l_sexo & "'"		
	l_sql = l_sql & "    ,observaciones  = '" & l_observaciones & "'"	
	
	
	l_sql = l_sql & "    ,fecha_ingreso  = " & l_fecha_ingreso
	l_sql = l_sql & "    ,fechanacimiento  = " & l_fechanacimiento
	l_sql = l_sql & "    ,nro_obra_social     = '" & l_nro_obra_social & "'"
	l_sql = l_sql & "    ,idciudad  = " & l_ciudad
	
	
	
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

