<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<%
'================================================================================
'Archivo		: borra_resultadosyarea_eva_00.asp
'Descripción	: Borra los resultados y areas correspondiente a....
'Autor			: 04-01-2005
'Fecha			: Leticia Amadio
'================================================================================

 dim l_sql
 dim l_cm
 dim l_rs

 dim l_evaevenro
 dim l_empleado

 l_evaevenro = Request.QueryString("evaevenro")
 l_empleado  = Request.QueryString("empleado")

 	' borra los datos de los resultados de las competencias
 l_sql ="DELETE FROM evaresultado" 
 l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
 l_sql = l_sql & " where evadetevldor.evldrnro=evaresultado.evldrnro "
 l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro
 l_sql = l_sql & " AND evacab.empleado = " & l_empleado & ")"
 set l_cm = Server.CreateObject("ADODB.Command") 
 ' response.write l_sql & "<br><br>"
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
  	' borra los datos del area de las competencias
l_sql = "DELETE FROM evaarea" 
l_sql = l_sql & " WHERE EXISTS (SELECT * FROM evadetevldor " 
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro" 
l_sql = l_sql & " where evadetevldor.evldrnro=evaarea.evldrnro " 
l_sql = l_sql & " AND evacab.evaevenro = " & l_evaevenro 
l_sql = l_sql & " AND evacab.empleado = " & l_empleado & ")" 
set l_cm = Server.CreateObject("ADODB.Command")
 'response.write l_sql & "<br>"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
 
%>

