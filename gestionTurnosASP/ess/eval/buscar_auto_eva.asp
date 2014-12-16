<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: buscar_auto_eva.asp
'Descripción	: Devuelve el mismo ternro :-)
'Autor			: CCRossi
'Fecha			: 18-05-2004
'Modificado		:
'================================================================================
Dim l_sql
dim l_cm

Dim l_ternro
Dim l_evldrnro
l_ternro	= request.QueryString("ternro")
l_evldrnro	= request.QueryString("evldrnro")

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "UPDATE evadetevldor SET "
l_sql = l_sql & " evaluador = " & l_ternro
l_sql = l_sql & " WHERE  evldrnro = "  & l_evldrnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
		
Set cn = Nothing
Set l_cm = Nothing
%>
<script>window.close();</script>


