<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 


Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_id
Dim l_descripcion



l_tipo        = request.querystring("tipo")
l_id          = request.Form("id")
l_descripcion = request.Form("descripcion")



	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO practicas"
		l_sql = l_sql & " (descripcion,empnro,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql = l_sql & " VALUES (" 
		l_sql = l_sql &  "'" & l_descripcion & "','" & session("empnro") & "','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
		
	else
		l_sql = "UPDATE practicas"
		l_sql = l_sql & " SET descripcion = '" & l_descripcion & "'"
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

