<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_03.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_expnro
Dim l_expdes

l_tipo 		= request.querystring("tipo")
l_expnro = request.Form("expnro")
l_expdes	= request.Form("expdes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_exportadora "
	l_sql = l_sql & " (expdes)"
	l_sql = l_sql & " VALUES ('" & l_expdes & "')"
else
	l_sql = "UPDATE buq_exportadora "
	l_sql = l_sql & " SET expdes = '" & l_expdes & "'"
	l_sql = l_sql & " WHERE expnro = " & l_expnro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

