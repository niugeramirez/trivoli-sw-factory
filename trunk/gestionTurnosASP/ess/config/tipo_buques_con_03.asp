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

Dim l_tipbuqnro
Dim l_tipbuqdes

l_tipo 		= request.querystring("tipo")
l_tipbuqnro = request.Form("tipbuqnro")
l_tipbuqdes	= request.Form("tipbuqdes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_tipobuque "
	l_sql = l_sql & " (tipbuqdes)"
	l_sql = l_sql & " VALUES ('" & l_tipbuqdes & "')"
else
	l_sql = "UPDATE buq_tipobuque "
	l_sql = l_sql & " SET tipbuqdes = '" & l_tipbuqdes & "'"
	l_sql = l_sql & " WHERE tipbuqnro = " & l_tipbuqnro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

