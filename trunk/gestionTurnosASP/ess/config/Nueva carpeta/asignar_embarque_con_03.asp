<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
on error goto 0
'Archivo: asignar_embarque_con_03.asp
'Descripción: ABM de Asignacion de Nros de camioneros de embarque
'Autor : Gustavo Manfrin
'Fecha: 20/09/2006

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_asiembnro
Dim l_embnro
Dim l_tarcod
Dim l_camnro
Dim l_camcha
Dim l_camaco
Dim l_tranro
Dim l_asiembobs

l_tipo 		= request.querystring("tipo")
l_asiembnro	= request.Form("asiembnro")
l_embnro	= request.Form("embnro")
l_tarcod	= request.Form("tarcod")
l_camnro	= request.Form("camnro")
l_camcha 	= request.Form("camcha")
l_camaco	= request.Form("camaco")
l_tranro	= request.Form("tranro")
l_asiembobs	= request.Form("asiembobs")


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_asiemb"
		l_sql = l_sql & " (embnro, tarcod, camnro, camcha, camaco, tranro, asiembobs)"
		l_sql = l_sql & " VALUES (" & l_embnro 
		l_sql = l_sql & ",'" & l_tarcod & "'," & l_camnro & ",'" & l_camcha & "','" & l_camaco & "'," & l_tranro
		l_sql = l_sql & ",'" & l_asiembobs & "'"
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE tkt_asiemb "
		l_sql = l_sql & " SET embnro = " & l_embnro
		l_sql = l_sql & ", tarcod = '" & l_tarcod & "'"
		l_sql = l_sql & ", camnro = " & l_camnro
		l_sql = l_sql & ", camcha = '" & l_camcha & "'"
		l_sql = l_sql & ", camaco = '" & l_camaco & "'"
		l_sql = l_sql & ", tranro = " & l_tranro
		l_sql = l_sql & ", asiembobs = '" & l_asiembobs & "'"
		l_sql = l_sql & " WHERE asiembnro = " & l_asiembnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

