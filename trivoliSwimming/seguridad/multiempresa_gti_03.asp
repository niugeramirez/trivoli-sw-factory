<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_mulnro
Dim l_mulnom
Dim l_multiple

function cambiafecha (actual)
  Dim auxi
  auxi  = mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4)
  cambiafecha = auxi
end function

l_tipo = request.querystring("tipo")
l_mulnro = request.Form("mulnro")
l_mulnom = request.Form("mulnom")
l_multiple = request.Form("multiple")

if l_multiple ="on" then
	l_multiple = -1
else
	l_multiple = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "insert into multiempresa "
	l_sql = l_sql & "(mulnom, multiple) "
	l_sql = l_sql & "values ('" & l_mulnom & "', " & l_multiple & ")"
else
	l_sql = "update multiempresa "
	l_sql = l_sql & "set mulnom = '" & l_mulnom & "', multiple = " & l_multiple & " where mulnro = " & l_mulnro
end if

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set cn = Nothing
Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'multiempresa_gti_01.asp';window.close();</script>"
%>
