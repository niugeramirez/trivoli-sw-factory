<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_cm
Dim l_sql


l_sql = request.querystring("consulta")

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
'Response.write "<script>alert('"&l_sql&"');</script>"
on error resume next
cmExecute l_cm, l_sql, 0
If err.Number <> 0 Then 
	Response.write(err.Number)
	Response.write(err.description)
else
	Set cn = Nothing
	Set l_cm = Nothing
	Response.write "<script> alert('Operación Realizada.'); opener.actualizar(); window.close();</script>"
end if	
%>
