<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: impresoras_con_03.asp
'Descripción: ABM de Impresoras
'Autor : Lisandro Moro
'Fecha: 26/09/2005
'Modificado: 

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_impnro
Dim l_impnom
Dim l_impnomcom
Dim l_impmat

l_tipo 		= request.querystring("tipo")
l_impnro	= request.Form("impnro")
l_impnom	= request.Form("impnom")
l_impnomcom	= request.Form("impnomcom")
l_impmat	= request.Form("impmat")
'Response.Write l_impmat
'Response.End

if l_impmat <> "on" then
	l_impmat = "0"
else
	l_impmat = "-1"
end if 

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_impresora"
		l_sql = l_sql & " (impnom, impnomcom, impmat )"
		l_sql = l_sql & " VALUES ('" & l_impnom & "','" & l_impnomcom & "','" & l_impmat & "')"
	else
		l_sql = "UPDATE tkt_impresora "
		l_sql = l_sql & " SET impnom = '" & l_impnom & "'"
		l_sql = l_sql & ", impnomcom = '" & l_impnomcom & "'"
		l_sql = l_sql & ", impmat = " & l_impmat 
  	    l_sql = l_sql & " WHERE impnro = " & l_impnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

