<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->

<%
'Archivo: camioneros_con_03.asp
'Descripción: Abm de camioneros
'Autor : Lisandro Moro
'Fecha: 15/02/2005

Dim l_tipo
Dim l_rs
Dim l_rs2
Dim l_cm
Dim l_sql

Dim l_camnro
Dim l_camcod
Dim l_camdes
Dim l_nrodoc
Dim l_tranro
Dim l_camcha
Dim l_camaco
Dim l_camcas
Dim l_camhab

on error goto 0 

l_camnro = request.QueryString("cabnro")
l_tipo	= request.QueryString("tipo")

'l_nrodoc = request.form("nrodoc")
l_camcod = request.form("camcod")
l_camdes = request.form("camdes")
l_tranro = request.form("tranro")
l_camcha = request.form("camcha")
l_camaco = request.form("camaco")
l_camcas = request.form("camcas")
l_camhab = request.form("camhab")

if l_camcas <> "-1" then
	l_camcas = 0
end if
if l_camhab <> "-1" then
	l_camhab = 0
end if

set l_cm = Server.CreateObject("ADODB.Command")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
'Realizo la validacio que no exista otro con el mismo codigo
l_sql = "SELECT camnro "
l_sql = l_sql & " FROM tkt_camionero "
l_sql = l_sql & " WHERE (camcod = " & l_camcod
l_sql = l_sql & " OR camdes = '" & l_camdes & "')"
if l_tipo = "M" then
	l_sql = l_sql & " AND camnro <> " & l_camnro
	'Response.Write "<script>alert(" & l_camnro & ");</script>"
end if

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	if l_tipo = "A" then 
		cn.BeginTrans
		l_cm.activeconnection = Cn
		
		l_sql = " INSERT INTO tkt_camionero "
		l_sql = l_sql & "( camcod, camdes, camcha, camaco, camcas, camhab ) "
		l_sql = l_sql & "VALUES (" & l_camcod & ",'" & l_camdes  & "','" & l_camcha & "','" & l_camaco & "'," & l_camcas & "," & l_camhab & ")"
		
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		'l_sql = fsql_seqvalue("next_id","tkt_camionero")
		'rsOpen l_rs2, cn, l_sql, 0
		'l_camnro = l_rs2("next_id")
		
		'l_sql = " INSERT INTO tkt_terdoc "
		'l_sql = l_sql & "( valnro, tipdocnro, tipternro, nrodoc ) "
		'l_sql = l_sql & "VALUES (" & l_camnro & ",5,2,'" & l_nrodoc & "')"
		
		'l_cm.CommandText = l_sql
		'cmExecute l_cm, l_sql, 0
		
		cn.CommitTrans
		'l_rs2.close
		set l_rs2 = nothing
	else
		cn.BeginTrans
		l_cm.activeconnection = Cn
		
		l_sql = "UPDATE tkt_camionero "
		l_sql = l_sql & " SET camcod = " & l_camcod
		l_sql = l_sql & " ,camdes = '" & l_camdes & "'"
		'l_sql = l_sql & " ,tranro = " & l_tranro
		l_sql = l_sql & " ,camcha = '" & l_camcha & "'"
		l_sql = l_sql & " ,camaco = '" & l_camaco & "'"
		l_sql = l_sql & " ,camcas = " & l_camcas
		l_sql = l_sql & " ,camhab = " & l_camhab
		l_sql = l_sql & " WHERE camnro = " & l_camnro
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		'l_sql = "UPDATE tkt_terdoc "
		'l_sql = l_sql & " SET nrodoc =  '" & l_nrodoc & "'"
		'l_sql = l_sql & " WHERE valnro = " & l_camnro
		'l_sql = l_sql & " AND tipternro = 2 AND tipdocnro = 5 "		
		'l_cm.CommandText = l_sql
		'cmExecute l_cm, l_sql, 0
		
		cn.CommitTrans
	end if
	cn.close 
	Set l_cm = Nothing
	Set cn = Nothing
else
	Response.Write "<script>alert('Ya existe el Código o la descripción del camionero.');</script>"
	'Response.Write "<script>window.close();</script>"
	Response.End
end if
'l_rs.Close
Set l_rs = Nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();window.close();</script>"
%>
