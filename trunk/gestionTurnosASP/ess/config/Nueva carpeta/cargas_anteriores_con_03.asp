<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: cargas_anteriores_con_03.asp
'Descripción: ABM de cargas anteriores
'Autor : Gustavo Manfrin
'Fecha: 07/08/2006

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_carconnro
Dim l_lugnro
Dim l_pronro

l_tipo 		= request.querystring("tipo")
'l_lugnro 	= request.querystring("lugnro")
l_lugnro 	= request.Form("lugnro")
l_pronro	= request.Form("pronro")
l_carconnro	= request.Form("carconnro")

	set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_cargasconf "
		l_sql = l_sql & " (pronro, lugdesnro)"
		l_sql = l_sql & " VALUES (" & l_pronro 
		l_sql = l_sql & "," & l_lugnro
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE tkt_cargasconf "
		l_sql = l_sql & " SET pronro = " & l_pronro
		l_sql = l_sql & ", lugnro = " & l_lugnro
		l_sql = l_sql & " WHERE  carconnro = " & l_carconnro
	end if
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	Set l_cm = Nothing
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

