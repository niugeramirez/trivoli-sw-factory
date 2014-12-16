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

Dim l_tipopenro
Dim l_tipopedes

l_tipo 		= request.querystring("tipo")
l_tipopenro = request.Form("tipopenro")
l_tipopedes	= request.Form("tipopedes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_tipoope "
	l_sql = l_sql & " (tipopedes)"
	l_sql = l_sql & " VALUES ('" & l_tipopedes & "')"
else
	l_sql = "UPDATE buq_tipoope "
	l_sql = l_sql & " SET tipopedes = '" & l_tipopedes & "'"
	l_sql = l_sql & " WHERE tipopenro = " & l_tipopenro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

