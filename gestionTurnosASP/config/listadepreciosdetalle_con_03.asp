<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_03.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_id
Dim l_idpractica
Dim l_precio

Dim l_idcab

l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")
l_idpractica  = request.Form("idpractica")
l_precio      = request.Form("precio2")

l_idcab = request.Form("idcab")


set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO listapreciosdetalle "
	l_sql = l_sql & " (idpractica, precio, idlistaprecioscabecera ,empnro,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES (" & l_idpractica & "," & l_precio & "," & l_idcab &"," & session("empnro") &",'"  &session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
else
	l_sql = "UPDATE listapreciosdetalle "
	l_sql = l_sql & " SET idpractica = " & l_idpractica 
	l_sql = l_sql & " , precio = " & l_precio
	'l_sql = l_sql & " , idlistaprecioscabecera = " & l_idcab 'Eugenio 03/04/2015 esto para mi no va, se me abortaba al agregar los campos who, creo que nunca se testeo
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

