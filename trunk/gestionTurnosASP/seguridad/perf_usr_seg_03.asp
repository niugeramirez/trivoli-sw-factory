<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: perf_usr_seg_03.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005

Dim l_tipo
Dim l_cm
Dim l_rs 'para realizar una consulta de control
Dim l_sql
Dim l_sql2
Dim l_perfnro
Dim l_perfnom
Dim l_perforden
Dim l_perftipo
Dim l_pol_nro


l_tipo = request.querystring("tipo")
l_perfnro = request.Form("perfnro")
l_perfnom = request.Form("perfnom")
'l_perforden = request.Form("perforden")
l_perftipo = request.Form("perftipo")
l_pol_nro	= request.Form("pol_nro")

if l_perftipo = "on" then
	l_perftipo = "-1"
else
	l_perftipo = "0"
end if

if l_pol_nro = "" then l_pol_nro = "null" end if

'controlamos que perfnom no exista

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "insert into perf_usr "
	l_sql = l_sql & "(perfnom, perftipo, pol_nro) "
	l_sql = l_sql & "values ('" & l_perfnom & "'," & l_perftipo & ", " & l_pol_nro & ")"
	l_sql2 = "select perfnro from perf_usr where perfnom = '" & l_perfnom & "' "
else
	l_sql = "update perf_usr "
	l_sql = l_sql & "set perfnom = '" & l_perfnom & "', perftipo = " & l_perftipo & ", pol_nro = " & l_pol_nro
	l_sql = l_sql & " where perfnro = " & l_perfnro
	l_sql2 = "select perfnro from perf_usr where perfnom = '" & l_perfnom & "' and perfnro <> " & l_perfnro	
end if
Set l_rs = Server.CreateObject("ADODB.RecordSet")	
l_cm.activeconnection = Cn
rsOpen l_rs, cn, l_sql2, 0
if l_rs.eof then
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	Set cn = Nothing
	Set l_cm = Nothing
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
else		
	Response.write "<script>alert('Existe un perfil con la descripción ingresada.');history.back();</script>"
end if
%>
