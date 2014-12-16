<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
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
		l_sql = "INSERT INTO obrassociales"
		l_sql = l_sql & " (descripcion)"
		l_sql = l_sql & " VALUES (" 
		l_sql = l_sql &  "'" & l_descripcion & "'"
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE obrassociales"
		l_sql = l_sql & " SET descripcion = '" & l_descripcion & "'"
  	    l_sql = l_sql & " WHERE id = " & l_id
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

