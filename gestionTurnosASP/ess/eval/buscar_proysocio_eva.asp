<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: buscar_proyrevisor_eva.asp
'Descripción	: Devuelve el ternro del revisor del proyecto
'Autor			: CCRossi
'Fecha			: 20-12-2004
'Modificado		:
'================================================================================
Dim l_sql
dim l_cm
dim l_rs

Dim l_ternro
Dim l_evldrnro
Dim l_evaluador
Dim l_evaproynro

l_ternro	= request.QueryString("ternro")
l_evldrnro	= request.QueryString("evldrnro")
l_evaproynro= request.QueryString("evaproynro")

l_evaluador=""

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT proysocio "
l_sql = l_sql & " FROM evaproyecto "
l_sql = l_sql & " INNER JOIN evaproyemp ON evaproyecto.evaproynro=evaproyemp.evaproynro "
l_sql = l_sql & "		 AND evaproyemp.ternro = " & l_ternro
l_sql = l_sql & " WHERE evaproyecto.evaproynro = " & l_evaproynro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_evaluador = l_rs("proysocio")
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
<script>window.close();</script>


