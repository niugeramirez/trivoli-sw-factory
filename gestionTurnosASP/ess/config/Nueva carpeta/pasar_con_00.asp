<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
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
	l_sql = "UPDATE tkt_transportista"
	l_sql = l_sql & " SET tracod = '" & tranro & "'"

	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

