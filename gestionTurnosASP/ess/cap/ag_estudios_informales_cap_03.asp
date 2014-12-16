<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: ag_estudios_informales_cap_03.asp
Descripción: Abm de Estudios Informales
Autor : Lisandro Moro	
Fecha: 29/03/2004
Modificacion:
-->
<% 
on error goto 0

Dim l_tipo
Dim l_rs
Dim l_cm
Dim l_sql
Dim l_ternro
Dim l_estinfdesabr
Dim l_estinffecha
Dim l_tipcurnro
Dim l_instnro
Dim l_estinfdesext
Dim l_estinfnro

l_tipo 			= request.querystring("tipo")
l_ternro		= l_ess_ternro

l_estinfnro		= request.Form("estinfnro")
l_estinfdesabr	= request.Form("estinfdesabr")
l_estinffecha  	= request.Form("estinffecha")
l_tipcurnro  	= request.Form("tipcurnro")
l_instnro 	 	= request.Form("instnro")
l_estinfdesext 	= request.Form("estinfdesext")

	set l_cm = Server.CreateObject("ADODB.Command")
	if l_tipo = "A" then 
		l_sql = "INSERT INTO cap_estinformal "
		l_sql = l_sql & "(ternro, estinfdesabr, estinffecha, tipcurnro, instnro, estinfdesext) "
		l_sql = l_sql & "VALUES (" & l_ternro & ",'" & l_estinfdesabr & "'," & cambiafecha(l_estinffecha ,"YMD",true) & "," & l_tipcurnro 
		l_sql = l_sql & "," & l_instnro & ",'" & l_estinfdesext & "') " 
	else
		l_sql = "UPDATE cap_estinformal "
		l_sql = l_sql & "SET estinffecha = " & cambiafecha(l_estinffecha ,"YMD",true)
		l_sql = l_sql & ",estinfdesabr = '" & l_estinfdesabr & "'"
		l_sql = l_sql & ",tipcurnro = " & l_tipcurnro
		l_sql = l_sql & ",instnro = " & l_instnro
		l_sql = l_sql & ",estinfdesext = '" & l_estinfdesext & "'"
		l_sql = l_sql & " WHERE estinfnro = " & l_estinfnro
	end if
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set l_cm = Nothing
	Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
