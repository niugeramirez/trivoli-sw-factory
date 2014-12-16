<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/inc/sqls.inc"-->
<!--#include virtual="/ticket/shared/inc/fecha.inc"-->

<%

'Archivo: terceros_documentos_con_03.asp
'Descripción: ABM de documentos asociados a tipos de terceros
'Autor : Alvaro bayon
'Fecha: 18/02/2005

on error goto 0 

Dim l_rs
Dim l_cm
Dim l_sql

Dim l_tipdocnro
Dim l_tipternro
Dim l_ternro

Dim l_fecvto
Dim l_oblig
Dim l_nrodoc
Dim l_nrodocant
Dim l_tipo


'l_nrodoc = request.form("nrodoc")
l_tipdocnro = request.form("tipdocnro")
l_tipternro = request.form("tipternro")
l_ternro = request.form("ternro")
l_fecvto = request.form("fecvto")
l_oblig = request.form("oblig")
l_nrodoc = request.form("nrodoc")
l_nrodocant = request.form("nrodocant")
l_tipo = request.form("tipo")

if l_fecvto <> "" then
	l_fecvto = cambiafecha(l_fecvto,"YMD",true)
else
	l_fecvto = "NULL"
end if

if l_nrodoc <> "" AND l_nrodocant = "" then
	l_tipo = "A"
else
	if l_nrodoc = "" AND l_nrodocant <> "" then
		l_tipo = "B"
	else
		if l_nrodoc = l_nrodocant then
			l_tipo = "X"
		else
			l_tipo = "M"
		end if
	end if
end if

select case l_tipo
	case "A"
		l_sql = " INSERT INTO tkt_terdoc "
		l_sql = l_sql & "( valnro, tipdocnro, tipternro, nrodoc, fecvto ) "
		l_sql = l_sql & "VALUES (" & l_ternro & "," & l_tipdocnro  & "," & l_tipternro & ",'" & l_nrodoc & "'," & l_fecvto & ")"
	case "M"
		l_sql = "UPDATE tkt_terdoc SET "
		l_sql = l_sql & " tipdocnro = " & l_tipdocnro
		l_sql = l_sql & " ,nrodoc = '" & l_nrodoc & "'"
		l_sql = l_sql & " ,fecvto = " & l_fecvto
		l_sql = l_sql & " WHERE valnro = " & l_ternro
		l_sql = l_sql & " AND tipternro = " & l_tipternro
	case "B"
		l_sql = " DELETE FROM tkt_terdoc "
		l_sql = l_sql & " WHERE valnro = " & l_ternro 
		l_sql = l_sql & " AND tipternro = " & l_tipternro
	case else
end select

if l_tipo <> "X" then
	set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if

'l_rs.Close
Set l_rs = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>
