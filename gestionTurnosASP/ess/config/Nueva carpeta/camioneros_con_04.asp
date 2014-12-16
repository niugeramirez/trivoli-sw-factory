<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'Archivo: camioneros_con_04.asp
'Descripción: Eliminar camioneros
'Autor : Lisandro Moro
'Fecha: 15/02/2005

Dim l_rs
Dim l_cm
Dim l_sql

Dim l_camnro

on error goto 0 

l_camnro = request.QueryString("cabnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Valido que no sea dato de sistema
l_sql = "SELECT camsis "
l_sql = l_sql & " FROM tkt_camionero "
l_sql = l_sql & " WHERE camnro = " & l_camnro
rsOpen l_rs, cn, l_sql, 0 
if l_rs("camsis") = -1 then
	l_rs.close
	set l_rs = nothing
	Response.Write "<script>alert('No se puede eliminar el Camionero por ser de Sistema.');window.close();</script>"
	Response.End
end if
l_rs.close
'Valido que no figure en alguna orden
l_sql = "SELECT camnro "
l_sql = l_sql & " FROM tkt_ord_cam "
l_sql = l_sql & " WHERE camnro = " & l_camnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_rs.close
	set l_rs = nothing
	Response.Write "<script>alert('No se puede eliminar el Camionero, \n está asociado a una Orden de Trabajo.');window.close();</script>"
	Response.End
end if
l_rs.close


set l_cm = Server.CreateObject("ADODB.Command")
cn.BeginTrans
	l_cm.activeconnection = Cn
	
	'recorro los documentos que tenga asociados
	l_sql = "SELECT tipdocnro "
	l_sql = l_sql & " FROM tkt_terdoc "
	l_sql = l_sql & " WHERE valnro = " & l_camnro
	l_sql = l_sql & " AND tipternro = 2 " 
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		do while not l_rs.eof
			l_sql = " DELETE FROM tkt_terdoc "
			l_sql = l_sql & " WHERE valnro = " & l_camnro 
			l_sql = l_sql & " AND tipdocnro = "  & l_rs("tipdocnro")
			l_sql = l_sql & " AND tipternro = 2 " 
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			l_rs.MoveNext
		loop
	end if
	l_rs.close
	
	'Elimino el camionero
	l_sql = " DELETE FROM tkt_camionero "
	l_sql = l_sql & " WHERE camnro = " & l_camnro 
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
cn.CommitTrans

set l_rs = nothing
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"
%>
