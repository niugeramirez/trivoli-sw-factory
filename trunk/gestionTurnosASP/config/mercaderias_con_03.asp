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

Dim l_mernro
Dim l_merdes
Dim l_tipmerdes
Dim l_merord

l_tipo 		 = request.querystring("tipo")
l_mernro	 = request.Form("mernro")
l_merdes	 = request.Form("merdes")
l_tipmerdes	 = request.Form("tipmerdes")
l_merord	 = request.Form("merord")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO buq_mercaderia "
	l_sql = l_sql & " (merdes, tipmerdes, merord)"
	l_sql = l_sql & " VALUES ('" & l_merdes & "','" & l_tipmerdes & "',"& l_merord &  "')"
else
	l_sql = "UPDATE buq_mercaderia "
	l_sql = l_sql & " SET merdes = '" & l_merdes & "'"
	l_sql = l_sql & " ,tipmerdes = '" & l_tipmerdes & "'"
	l_sql = l_sql & " ,merord = " & l_merord
	l_sql = l_sql & " WHERE mernro = " & l_mernro
end if
'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

