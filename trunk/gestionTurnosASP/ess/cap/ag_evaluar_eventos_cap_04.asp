
<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->

<!--
Archivo: 		ag_evaluar_eventos_cap_04.asp
Descripción: 	Graba las Respuestas de los Formularios
Autor : 		Raul Chinestra
Fecha: 			25/06/2007
-->
<% 

on error goto 0

'Datos del formulario
Dim l_ttesnro
Dim l_ttesdesabr
Dim l_ttesdesext
Dim l_ttespond
Dim l_ttessistema
Dim l_ttescosto
Dim l_tesnro

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_cm

dim item
Dim l_prenro
Dim l_pretipo

Dim l_fornro
Dim l_ternro

l_ternro = l_ess_ternro

l_fornro = Request.queryString("fornro")
l_tesnro = Request.queryString("tesnro")
l_prenro = Request.form("prenro")
l_pretipo = Request.form("pretipo")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn

'----------------------------------------------------------------------------------------------------------------
' Busco el Tipo de Test Asociado al Formulario
'----------------------------------------------------------------------------------------------------------------
l_sql = "SELECT ttesnro " 
l_sql = l_sql & " FROM pos_formulario "
l_sql = l_sql & " WHERE pos_formulario.fornro = " & l_fornro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_ttesnro	= l_rs("ttesnro")
end if
l_rs.close


'----------------------------------------------------------------------------------------------------------------
' Elimino las Respuestas ingresadas previmante
'----------------------------------------------------------------------------------------------------------------
l_sql = "SELECT * " 
l_sql = l_sql & " FROM pos_pregunta "
l_sql = l_sql & " WHERE pos_pregunta.fornro = " & l_fornro
l_sql = l_sql & " AND prenro = " & l_prenro
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	l_sql = "DELETE FROM pos_respuesta "
	l_sql = l_sql & " WHERE tesnro = " & l_tesnro & " AND prenro = " & l_rs("prenro") 
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	l_rs.MoveNext
loop


'/* luego ciclo por todas las respuestas que vengan por el form */
dim l_arreglo
dim l_request
dim a
Dim l_rta
Dim l_res
l_request = Request.form

'Response.write "<br>" & Request.form & "<br>"
'l_arreglo = split(l_request, "&")
'for a = 0 to UBound(l_arreglo)'
'	l_pretipo = Split(l_arreglo(a),"=")
'	l_rta = Split(l_arreglo(a + 1),"=")
'	l_prenro = l_rta(0)
'	l_res = Request.form(a + 2)
'
	l_res = Request.form(l_prenro)
	l_sql = "INSERT INTO pos_respuesta "
	if CInt(l_pretipo) <> 0 then
		l_sql = l_sql & "(tesnro, prenro, resdes) "
		l_sql = l_sql & " VALUES (" & l_tesnro & ","& l_prenro & ",'" & l_res & "')"
	else
		l_sql = l_sql & "(tesnro, prenro, resval) "
		l_sql = l_sql & " VALUES (" & l_tesnro & ","& l_prenro & "," & l_res & ")"
	end if
	
'	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
'	'if l_pretipo(1) = -1 AND l_res <> "" then
	cmExecute l_cm, l_sql, 0
'	'end if
'
'	a = a + 1



	Response.write l_sql & "<br>"
'	response.write l_arreglo(a) & "<br>"
'Next
'response.End()
'response.write "<script>alert('Operación Realizada.');window.opener.close();window.close();</script>"

'---------------------------------------------------------------------------------
set l_rs = nothing
%>
</body>
</html>
