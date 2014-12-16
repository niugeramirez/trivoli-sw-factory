<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: depositos_con_03.asp
'Descripción: ABM de Tipos de Mermas para rubros
'Autor : Alvaro Bayon
'Fecha: 16/02/2005

Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_tipmernro
Dim l_forcal
Dim l_rubnro
Dim l_lugnro
Dim l_tipmer

l_tipo 		= request.querystring("tipo")
l_tipmernro = request.Form("tipmernro")
l_forcal	= request.Form("forcal")
l_lugnro 	= request.Form("lugnro")
l_rubnro	= request.Form("rubnro")
l_tipmer	= request.Form("tipmer")

'if len(l_forcal)=0 then
'	l_forcal = "null"
'else
'	l_forcal = "'" & l_forcal & "'"
'end if
'
'	set l_cm = Server.CreateObject("ADODB.Command")
'	if l_tipo = "A" then 
'		l_sql = "INSERT INTO tkt_tipomerma"
'		l_sql = l_sql & " (forcal, lugnro, tipmer, rubnro)"
'	'	l_sql = l_sql & " VALUES (" & l_forcal 
'		l_sql = l_sql & "," & l_lugnro & ",'" & l_tipmer & "'," & l_rubnro
''		l_sql = l_sql & ")"
''	else
'		l_sql = "UPDATE tkt_tipomerma"
'		l_sql = l_sql & " SET forcal = " & l_forcal 
'		l_sql = l_sql & ", lugnro = " & l_lugnro
'		l_sql = l_sql & ", tipmer = '" & l_tipmer & "'"
'		l_sql = l_sql & ", rubnro = " & l_rubnro
'		l_sql = l_sql & " WHERE tipmernro = " & l_tipmernro
'	end if
'	'response.write l_sql & "<br>"
'	l_cm.activeconnection = Cn
''	l_cm.CommandText = l_sql
'	cmExecute l_cm, l_sql, 0
'	Set l_cm = Nothing

	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
	
%>

