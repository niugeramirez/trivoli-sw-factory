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

Dim l_desnro
Dim l_desdes

l_tipo 		 = request.querystring("tipo")
l_desnro	 = request.Form("desnro")
l_desdes	 = request.Form("desdes")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_destino "
	l_sql = l_sql & " (desdes)"
	l_sql = l_sql & " VALUES ('" & l_desdes & "')"
else
	l_sql = "UPDATE buq_destino "
	l_sql = l_sql & " SET desdes = '" & l_desdes & "'"
	l_sql = l_sql & " WHERE desnro = " & l_desnro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

