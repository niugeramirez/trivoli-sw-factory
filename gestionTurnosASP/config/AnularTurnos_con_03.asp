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
dim l_motivo
dim l_estado
dim l_cantturnossimult  
dim l_cantsobreturnos     



l_tipo 		               = request.querystring("tipo")
l_id                       = request.Form("id")
l_motivo                   = request.Form("motivo")

' l_domicilio      = request.Form("domicilio")
'l_idobrasocial      = request.Form("legape")




if l_tipo = "B" then
	l_estado = "ANULADO"
else
	l_estado = "ACTIVO"
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

l_sql = "UPDATE calendarios "
l_sql = l_sql & " SET motivo    = '" & l_motivo & "'"
l_sql = l_sql & "    ,estado    = '" & l_estado & "' "

l_sql = l_sql & " WHERE id = " & l_id

response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

