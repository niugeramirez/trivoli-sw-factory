<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'Archivo	: tieneobjetivos_eva_ag_02.asp
'Descripción: poner tieneobj en -1
'Autor		: CCRossi
'Fecha		: 26-11-2004
'Modificacion: 
'-------------------------------------------------------------------------------
 Dim l_cm
 Dim l_sql
 dim l_rs 

'parametros de entrada
 dim l_evaluador
 dim l_lista
 dim l_listainicial
 dim l_arreglo
   
l_evaluador = Request.QueryString("evaluador") ' viene el ternro del logeado
l_lista  = Request.QueryString("lista")
l_listainicial  = Request.QueryString("listainicial")

if trim(l_lista)="" then
l_lista="0"
end if
' ------------------------------------------------------------------------------------
'											BODY 
' ------------------------------------------------------------------------------------
l_arreglo= Split(l_lista,",")


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evacab.evacabnro  "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
l_sql = l_sql & "		AND  evadetevldor.evaluador = " & l_evaluador
l_sql = l_sql & " WHERE evacab.empleado  IN (" & l_lista &")"
rsOpen l_rs, cn, l_sql, 0 
do while not l_rs.eof 

	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE  evacab SET "
	l_sql = l_sql & " tieneobj=  -1 "
	l_sql = l_sql & " WHERE evacabnro="& l_rs("evacabnro")
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_rs.MoveNext
		
loop	
l_rs.Close
set l_rs=nothing


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evacab.evacabnro  "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
l_sql = l_sql & "		AND  evadetevldor.evaluador = " & l_evaluador
l_sql = l_sql & " WHERE evacab.empleado  NOT IN (" & l_lista &")"
l_sql = l_sql & "  AND  evacab.empleado  IN (" & l_listainicial &")"
rsOpen l_rs, cn, l_sql, 0 
do while not l_rs.eof 

	set l_cm = Server.CreateObject("ADODB.Command")
	l_sql = "UPDATE  evacab SET "
	l_sql = l_sql & " tieneobj=  0 "
	l_sql = l_sql & " WHERE evacabnro="& l_rs("evacabnro")
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_rs.MoveNext
		
loop	
l_rs.Close
set l_rs=nothing

cn.close
Set cn = Nothing

response.write "<script>window.returnValue='0';window.close();</script>"
%>