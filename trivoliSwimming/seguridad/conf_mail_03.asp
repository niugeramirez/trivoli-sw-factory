<%  Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%
'Archivo        : conf_mil_03.asp
'Descripcion    : Modulo que se encarga de admin. los servidores de mail
'Creador        : Lisandro Moro
'Fecha Creacion : 08/03/2005
'Modificacion   :

on error goto 0

' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_cfgemailnro
Dim l_cfgemailhost
Dim l_cfgemaildesc
Dim l_cfgemailfrom
Dim l_cfgemailest
Dim l_cfgemailport

Dim l_cfgemailhostant
Dim l_cfgemaildescant
Dim l_cfgemailfromant
Dim l_cfgemailestant
Dim l_cfgemailportant

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_cfgemailnro  = request.Form("cfgemailnro")
l_cfgemailhost = request.Form("cfgemailhost")
l_cfgemaildesc = request.Form("cfgemaildesc")
l_cfgemailfrom = request.Form("cfgemailfrom")
l_cfgemailest  = request.Form("cfgemailest")
l_cfgemailport = request.Form("cfgemailport")

l_cfgemailhostant = request.Form("cfgemailhostant")
l_cfgemaildescant = request.Form("cfgemaildescant")
l_cfgemailfromant = request.Form("cfgemailfromant")
l_cfgemailestant  = request.Form("cfgemailestant")
l_cfgemailportant = request.Form("cfgemailportant")

if l_cfgemailest = "on" then
	l_cfgemailest = -1
else
	l_cfgemailest = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")

cn.beginTrans

if (CInt(l_cfgemailest) = -1) AND  (l_cfgemailestant <> l_cfgemailest) then
	l_sql = "UPDATE conf_email SET "
	l_sql = l_sql & " cfgemailest = 0 " 

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	'response.write l_sql
	cmExecute l_cm, l_sql, 0	
end if

if l_tipo = "A" then 

	  l_sql = "INSERT INTO conf_email "
	  l_sql = l_sql & "(cfgemailhost,cfgemaildesc,cfgemailfrom,cfgemailest,cfgemailport) "
	  l_sql = l_sql & " VALUES ('" 
	  l_sql = l_sql & l_cfgemailhost & "','"
	  l_sql = l_sql & l_cfgemaildesc  & "','"
	  l_sql = l_sql & l_cfgemailfrom  & "',"
	  l_sql = l_sql & l_cfgemailest  & ","
	  l_sql = l_sql & l_cfgemailport  & ")"

	  l_cm.activeconnection = Cn
	  l_cm.CommandText = l_sql
	  cmExecute l_cm, l_sql, 0	

else

		l_sql = "UPDATE conf_email SET "
		l_sql = l_sql & " cfgemaildesc  = '" & l_cfgemaildesc & "',"
		l_sql = l_sql & " cfgemailhost  = '" & l_cfgemailhost & "',"
		l_sql = l_sql & " cfgemailfrom  = '" & l_cfgemailfrom & "',"
		l_sql = l_sql & " cfgemailest   =  " & l_cfgemailest  & ","
		l_sql = l_sql & " cfgemailport  =  " & l_cfgemailport & " "
		l_sql = l_sql & " WHERE cfgemailnro = "  & l_cfgemailnro

		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0	
end if

cn.CommitTrans

Response.write "<script>alert('Operación Realizada.');window.opener.opener.ifrm.location.reload();window.opener.close();window.close();</script>"
%>
