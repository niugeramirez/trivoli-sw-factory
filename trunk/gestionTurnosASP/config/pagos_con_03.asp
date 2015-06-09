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
Dim l_idmediodepago
Dim l_fecha
Dim l_idobrasocial
Dim l_idpracticarealizada
Dim l_nro
Dim l_importe

l_tipo 		  = request.querystring("tipo")
l_id 	      = request.Form("id")
l_idmediodepago	  = request.Form("idmediodepago")
l_fecha       = request.Form("fecha")
l_idobrasocial = request.Form("idobrasocial")

if l_idobrasocial = "" then l_idobrasocial = 0 end if

l_idpracticarealizada = request.Form("idpracticarealizada")
l_nro      = request.Form("nro")
l_importe      = request.Form("importe2")


if len(l_fecha) = 0 then
	l_fecha = "null"
else 
	l_fecha = cambiafecha(l_fecha,"YMD",true)	
end if 




set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO pagos "
	l_sql = l_sql & " (idmediodepago, fecha, idpracticarealizada, idobrasocial, nro, importe ,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES (" & l_idmediodepago & "," & l_fecha & "," & l_idpracticarealizada & "," & l_idobrasocial & ",'" & l_nro & "'," & l_importe &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"	
	
else
	l_sql = "UPDATE pagos "
	l_sql = l_sql & " SET idmediodepago = " & l_idmediodepago
	l_sql = l_sql & " , fecha = " & l_fecha 
	l_sql = l_sql & " , idpracticarealizada = " & l_idpracticarealizada
	l_sql = l_sql & " , idobrasocial = " & l_idobrasocial
	l_sql = l_sql & " , nro = '" & l_nro & "'"
	l_sql = l_sql & " , importe = " & l_importe		
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.parent.opener.ifrm.location.reload();window.parent.opener.close();window.parent.close();</script>"
%>

