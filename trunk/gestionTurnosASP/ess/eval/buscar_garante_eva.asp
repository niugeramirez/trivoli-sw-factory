<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: buscar_garante_eva.asp
'Descripción	: Devuelve el ternro del garante
'Autor			: CCRossi
'Fecha			: 21-02-2005
'Modificado		:
'================================================================================
' debe existir const ctenroarea	= 10  ' tenro de area
' debe existir const ctenrogarante	= 12  ' tenro ded tipo estructura garante

Dim l_sql
dim l_cm
dim l_rs

Dim l_ternro
Dim l_evldrnro
Dim l_evaluador

l_ternro	= request.QueryString("ternro")
l_evldrnro	= request.QueryString("evldrnro")

l_evaluador=""

'Buscar el área del empelado:
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estrnro "
l_sql = l_sql & " FROM his_estructura "
l_sql = l_sql & " WHERE his_estructura.ternro = " & l_ternro
l_sql = l_sql & "   AND his_estructura.htethasta IS NULL " 
l_sql = l_sql & "   AND his_estructura.tenro ="  & ctenroarea
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then

	'Buscar un garante con esta area
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT his_estructura.ternro "
	l_sql = l_sql & " FROM his_estructura "
	l_sql = l_sql & " INNER JOIN his_estructura area ON his_estructura.ternro = area.ternro "
	l_sql = l_sql & "		 AND area.tenro   = " & ctenroarea
	l_sql = l_sql & "		 AND area.estrnro = " & l_rs("estrnro")
	l_sql = l_sql & " WHERE his_estructura.tenro = " & ctenrogarante
	l_sql = l_sql & " WHERE his_estructura.htethasta IS NULL " 
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.EOF then
		l_evaluador = l_rs("ternro")
		l_rs.close
		set l_rs=nothing
	
		if trim(l_evaluador)<>"" AND not isnull(l_evaluador) then
			set l_cm = Server.CreateObject("ADODB.Command")
			l_sql = "UPDATE evadetevldor SET "
			l_sql = l_sql & " evaluador = " & l_evaluador	
			l_sql = l_sql & " WHERE  evldrnro = "  & l_evldrnro
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if	
end if		

Set cn = Nothing
%>
<script>window.close();</script>


