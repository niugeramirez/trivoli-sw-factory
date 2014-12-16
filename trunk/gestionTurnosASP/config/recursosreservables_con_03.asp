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
dim l_descripcion
dim l_cantturnossimult  
dim l_cantsobreturnos     



l_tipo 		               = request.querystring("tipo")
l_id                       = request.Form("id")
l_descripcion              = request.Form("descripcion")
l_cantturnossimult         = request.Form("cantturnossimult")
l_cantsobreturnos          = request.Form("cantsobreturnos")
' l_domicilio      = request.Form("domicilio")
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
	l_sql = "INSERT INTO recursosreservables  "
	l_sql = l_sql & " (descripcion, cantturnossimult, cantsobreturnos)"
	l_sql = l_sql & " VALUES ('" & l_descripcion & "'," & l_cantturnossimult & "," & l_cantsobreturnos & ")"
else
	l_sql = "UPDATE recursosreservables "
	l_sql = l_sql & " SET descripcion    = '" & l_descripcion & "'"
	l_sql = l_sql & "    ,cantturnossimult    = " & l_cantturnossimult & ""
	l_sql = l_sql & "    ,cantsobreturnos    =    " & l_cantsobreturnos & ""
	l_sql = l_sql & " WHERE id = " & l_id
end if
response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

