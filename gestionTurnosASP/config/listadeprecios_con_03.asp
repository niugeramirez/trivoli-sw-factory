<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 
'Archivo: companies_con_03.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_id
Dim l_titulo
Dim l_fecha
Dim l_flag_activo
Dim l_idobrasocial

l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")
l_titulo	  = request.Form("titulo")
l_fecha       = request.Form("fecha")
l_flag_activo = request.Form("activo")
l_idobrasocial = request.Form("idobrasocial")

if len(l_fecha) = 0 then
	l_fecha = "null"
else 
	l_fecha = cambiafecha(l_fecha,"YMD",true)	
end if 




set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO listaprecioscabecera "
	l_sql = l_sql & " (titulo, fecha, idobrasocial, flag_activo ,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES ('" & l_titulo & "'," & l_fecha & "," & l_idobrasocial & "," & l_flag_activo &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
else
	l_sql = "UPDATE listaprecioscabecera "
	l_sql = l_sql & " SET titulo = '" & l_titulo & "'"
	l_sql = l_sql & " , fecha = " & l_fecha 
	'l_sql = l_sql & " , idobrasocial = " & l_idobrasocial 'Eugenio 03/04/2015 esto para mi no va, se me abortaba al agregar los campos who, creo que nunca se testeo
	l_sql = l_sql & " , flag_activo = " & l_flag_activo
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

