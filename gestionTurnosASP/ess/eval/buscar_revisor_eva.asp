<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: buscar_revisor_eva.asp
'Descripción	: Devuelve el ternro del revisor
'Autor			: CCRossi
'Fecha			: 19-05-2004
'Modificado		:
'================================================================================
Dim l_sql
dim l_cm
dim l_rs

Dim l_ternro
Dim l_evldrnro
Dim l_evaluador

l_ternro	= request.QueryString("ternro")
l_evldrnro	= request.QueryString("evldrnro")

l_evaluador=""
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado.ternro, empleado.empreporta "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " LEFT JOIN empleado ON empleado.ternro = empleado.empreporta"
l_sql = l_sql & " WHERE empleado.ternro = " & l_ternro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_evaluador = l_rs("ternro")
	l_rs.close
	set l_rs=nothing
	
	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE evadetevldor SET "
	if trim(l_evaluador)=""	or isnull(l_evaluador) then
		l_sql = l_sql & " evaluador = NULL " 
	else
		l_sql = l_sql & " evaluador = " & l_evaluador
	end if
	l_sql = l_sql & " WHERE  evldrnro = "  & l_evldrnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql 
	cmExecute l_cm, l_sql, 0 
end if		
Set cn = Nothing
Set l_cm = Nothing
%>
<script> window.close();</script>


