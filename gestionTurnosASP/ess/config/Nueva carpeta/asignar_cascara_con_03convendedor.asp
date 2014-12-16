<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
on error goto 0
'Archivo: asignar_cascara_con_03.asp
'Descripción: ABM de Asignacion de Nros de Cáscara
'Autor : Raul Chinestra
'Fecha: 09/05/2005

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_asicasnro
Dim l_ordnro
Dim l_tarnro
Dim l_camnro
Dim l_camcha
Dim l_camaco
Dim l_tranro
Dim l_vennro
Dim l_cornro
Dim l_entnro
Dim l_desnro
Dim l_deporinro
Dim l_depdesnro

l_tipo 		= request.querystring("tipo")
l_asicasnro	= request.Form("asicasnro")
l_ordnro	= request.Form("ordnro")
l_tarnro	= request.Form("tarnro")
l_camnro	= request.Form("camnro")
l_camcha 	= request.Form("camcha")
l_camaco	= request.Form("camaco")
l_tranro	= request.Form("tranro")
l_vennro	= request.Form("vennro")
l_cornro	= request.Form("cornro")
l_entnro	= request.Form("entnro")
l_desnro	= request.Form("desnro")
l_deporinro	= request.Form("deporinro")
l_depdesnro	= request.Form("depdesnro")


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_asicas"
		l_sql = l_sql & " (ordnro, tarnro, camnro, camcha, camaco, tranro, vennro, cornro, entnro, desnro, deporinro, depdesnro)"
		l_sql = l_sql & " VALUES (" & l_ordnro 
		l_sql = l_sql & ",'" & l_tarnro & "'," & l_camnro & ",'" & l_camcha & "','" & l_camaco & "'," & l_tranro
		l_sql = l_sql & "," & l_vennro & "," & l_cornro & "," & l_entnro & "," & l_desnro & "," & l_deporinro & "," & l_depdesnro
		l_sql = l_sql & ")"
	else
		l_sql = "UPDATE tkt_asicas"
		l_sql = l_sql & " SET ordnro = " & l_ordnro
		l_sql = l_sql & ", tarnro = '" & l_tarnro & "'"
		l_sql = l_sql & ", camnro = " & l_camnro
		l_sql = l_sql & ", camcha = '" & l_camcha & "'"
		l_sql = l_sql & ", camaco = '" & l_camaco & "'"
		l_sql = l_sql & ", tranro = " & l_tranro
		l_sql = l_sql & ", vennro = " & l_vennro
		l_sql = l_sql & ", cornro = " & l_cornro
		l_sql = l_sql & ", entnro = " & l_entnro
		l_sql = l_sql & ", desnro = " & l_desnro
		l_sql = l_sql & ", deporinro = " & l_deporinro
		l_sql = l_sql & ", depdesnro = " & l_depdesnro
		l_sql = l_sql & " WHERE asicasnro = " & l_asicasnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

