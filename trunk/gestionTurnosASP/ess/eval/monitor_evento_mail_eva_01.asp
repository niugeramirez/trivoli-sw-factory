<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/ess/ess/shared/inc/mails.inc"-->
<%
'===========================================================================================
'Archivo	: monitor_evento_mail_eva_01.asp
'Descripción: envio mail para la lista del monitor
'Autor		: CCRossi
'Fecha		: 26-07-2004
'Modificado :
'===========================================================================================

'Parametro de entrada
 Dim l_listamail
 Dim l_Body
 Dim l_cuerpo
 Dim l_Subject
 
'Locales
 Dim l_sql
 Dim l_rs

 Dim l_empemail 
 Dim l_usrname 
 Dim l_usrmail 
 Dim l_nombre
 
l_listamail = request("listamail")
l_cuerpo	= request("Body")
l_Subject	= request("Subject")


if trim(l_listamail)="" then%>
	<script>
		window.close();
	</script>
<%
else
	l_listamail = "0" & l_listamail
end if
			

Dim arch
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  distinct ternro,terape, terape2, ternom, ternom2, empemail "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE empleado.ternro IN (" & l_listamail & ")"
l_sql = l_sql & " AND   empleado.empemail IS NOT NULL "
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.EOF then
 l_nombre=""
  do while not l_rs.eof 
	l_empemail = l_rs("empemail")
	if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
		l_nombre = l_rs("ternom")
	end if
	if trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & " " & l_rs("ternom2")
	end if
	if trim(l_rs("terape2"))<>"" then
		l_nombre = l_nombre & " "& l_rs("terape2") 
	end if
	
	l_Body = l_nombre & ", " & l_cuerpo
	
	
	if trim(l_empemail)<>"" then
	    'response.write ("<script>alert('"&l_empemail&"')</script>")
		'generarMail "",l_subject,l_Body,l_empemail
		'Envio el email
		
		enviarMail fs,"",l_subject,l_Body,l_empemail
	
	end if
	l_nombre=""
	l_rs.MoveNext
 loop	
 l_rs.close
 set l_rs=nothing
 Response.Write("<script>alert('Mails enviados a los Evaluadores de la lista.');window.close();</script>")
else
 Response.Write("<script>window.close();</script>")
end if
%>
