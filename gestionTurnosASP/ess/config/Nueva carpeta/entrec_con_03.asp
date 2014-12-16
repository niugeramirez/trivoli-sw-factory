<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: entrec_con_03.asp
'Descripción: Abm de Entregadores y recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_entnro
Dim l_entdes
Dim l_entrol
Dim l_entcod

l_tipo 		= request.querystring("tipo")
l_entnro 	= request.Form("entnro")
l_entdes	= request.Form("entdes")
l_entcod 	= request.Form("entcod")
l_entrol	= request.Form("entrol")

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO tkt_entrec"
		l_sql = l_sql & " (entdes, entcod, entrol)"
		l_sql = l_sql & " VALUES ('" & l_entdes 
		l_sql = l_sql & "','" & l_entcod & "','" & l_entrol
		l_sql = l_sql & "')"
	else
		l_sql = "UPDATE tkt_entrec"
		l_sql = l_sql & " SET entdes = '" & l_entdes & "'"
		l_sql = l_sql & ", entcod = '" & l_entcod & "'"
		l_sql = l_sql & ", entrol = '" & l_entrol & "'"
		l_sql = l_sql & " WHERE entnro = " & l_entnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

