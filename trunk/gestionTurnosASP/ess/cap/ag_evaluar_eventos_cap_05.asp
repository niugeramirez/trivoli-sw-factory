
<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->

<!--
Archivo: 		ag_evaluar_eventos_cap_05.asp
Descripción: 	Cierra el formulario Formularios
Autor : 		LIsandro Moro
Fecha: 			06/07/2007
-->
<% 

on error goto 0

'Datos del formulario
Dim l_tesnro
Dim l_testie

'ADO
Dim l_sql
Dim l_rs
Dim l_cm

l_tesnro = Request.queryString("tesnro")
l_testie = Request.queryString("testie")

set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn

'----------------------------------------------------------------------------------------------------------------
' Elimino las Respuestas ingresadas previmante
'----------------------------------------------------------------------------------------------------------------
	l_sql = "UPDATE test "
	l_sql = l_sql & " SET tesfin = -1 "
	l_sql = l_sql & ", testie = '" & split(l_testie,":")(1) & ":" & split(l_testie,":")(2) & "'"
	l_sql = l_sql & " WHERE tesnro = " & l_tesnro
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

response.write "<script>alert('Operación Realizada.');window.parent.location.reload();window.close();</script>"

'---------------------------------------------------------------------------------
set l_cm = nothing
set l_rs = nothing
%>
</body>
</html>
