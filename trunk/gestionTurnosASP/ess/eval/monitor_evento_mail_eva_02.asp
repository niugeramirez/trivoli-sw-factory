<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'===========================================================================================
'Archivo	: monitor_evento_mail_eva_02.asp
'Descripción: envio mail para cada empemail
'Autor		: CCRossi
'Fecha		: 26-07-2004
'Modificado :
'===========================================================================================

'Parametro de entrada
 Dim l_empemail
 Dim l_Body
 Dim l_Subject
 
'Locales
 Dim l_sql
 Dim l_rs
 Dim l_usrname 
 Dim l_usrmail 
 
l_empemail  = request("empemail")
l_Body		= request("Body")
l_Subject	= request("Subject")
l_usrname	= "RRHH"
l_usrmail	= "RRHH"

if trim(l_empemail)="" then%>
	<script>
		window.close();
	</script>
<%end if

'Response.Write("<script>alert('"&l_empemail&"');</script>")
'Response.Write("<script>alert('"&l_Body&"');</script>")
'Response.Write("<script>alert('"&l_Subject&"');</script>")
'Response.End
%>
