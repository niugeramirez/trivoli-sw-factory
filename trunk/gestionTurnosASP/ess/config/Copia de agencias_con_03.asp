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

Dim l_agenro
Dim l_agedes

l_tipo 		= request.querystring("tipo")
l_agenro 	= request.Form("agenro")
l_agedes	= request.Form("agedes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_agencia "
	l_sql = l_sql & " (agedes)"
	l_sql = l_sql & " VALUES ('" & l_agedes & "')"
else
	l_sql = "UPDATE buq_agencia "
	l_sql = l_sql & " SET agedes = '" & l_agedes & "'"
	l_sql = l_sql & " WHERE agenro = " & l_agenro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

