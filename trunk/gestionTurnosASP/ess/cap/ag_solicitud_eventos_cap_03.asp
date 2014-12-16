<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_00.asp
Descripción: Abm de Solicitud de Eventos
Autor : Raul CHinestra
Fecha: 30/03/2004
-->
<% 

on error goto 0

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql
Dim l_solnro
Dim l_soldesabr
Dim l_soldesext
Dim l_soldurdias
Dim l_solfec

Dim l_ternro

l_tipo 			= request.querystring("tipo")
l_solnro 		= request.Form("solnro")
l_soldesabr 	= request.Form("soldesabr")
l_solfec	 	= request.Form("solfec")
l_soldurdias 	= request.Form("soldurdias")
l_soldesext 	= request.Form("soldesext")

l_ternro	 	= l_ess_ternro

if l_soldurdias = "" then
	l_soldurdias = 0
end if 


	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO cap_solicitud "
		l_sql = l_sql & "(soldesabr, soldesext, soldurdias, ternro,solfec) "
		l_sql = l_sql & "VALUES ('" & l_soldesabr & "'"
		l_sql = l_sql & ",'" & l_soldesext & "'"
		l_sql = l_sql & "," & l_soldurdias 
		l_sql = l_sql & "," & l_ternro & "," & cambiafecha(l_solfec,"YMD",true)  & ")"		
	else
		l_sql = "UPDATE cap_solicitud "
		l_sql = l_sql & "SET soldesabr = '" & l_soldesabr & "'"
		l_sql = l_sql & ",soldurdias = " & l_soldurdias
		l_sql = l_sql & ",soldesext = '" & l_soldesext & "'"
		l_sql = l_sql & ",solfec = " & cambiafecha(l_solfec,"YMD",true)  
		l_sql = l_sql & " WHERE solnro = " & l_solnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing
	cn.close
	Set cn = nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
