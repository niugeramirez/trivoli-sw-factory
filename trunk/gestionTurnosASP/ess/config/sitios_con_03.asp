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

Dim l_sitnro
Dim l_sitdes

l_tipo 		= request.querystring("tipo")
l_sitnro = request.Form("sitnro")
l_sitdes	= request.Form("sitdes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_sitio "
	l_sql = l_sql & " (sitdes)"
	l_sql = l_sql & " VALUES ('" & l_sitdes & "')"
else
	l_sql = "UPDATE buq_sitio "
	l_sql = l_sql & " SET sitdes = '" & l_sitdes & "'"
	l_sql = l_sql & " WHERE sitnro = " & l_sitnro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

