<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<%
'---------------------------------------------------------------------------------
'Archivo	: evaluadores_eva_01.asp
'Descripción: graba los ternro cagados a mano
'Autor		: 
'Fecha		: 
'----------------------------------------------------------------------------------


  Dim l_cm
  Dim l_sql
  Dim l_rs
  
  
' variables
' parametros de entrada
  Dim l_lista
  Dim l_empleado
  Dim l_evaevenro 
  
'uso local
  Dim l_evaluador
  Dim l_empleg
  Dim i
  dim l_evatevnro 
  
  Dim l_arreglo
  
  l_lista		= Request.QueryString("lista")
  l_empleado	= Request.QueryString("ternro")
  l_evaevenro	= request.QueryString("evaevenro")

' -------------------------------------------------------------------
'							BODY
' -------------------------------------------------------------------

'cn.BeginTrans

l_arreglo = split(l_lista,",")  
i=0
do while  i<= Ubound(l_arreglo)-1

	l_evatevnro = l_arreglo(i)
	l_empleg = l_arreglo(i+1)

	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = " SELECT ternro FROM empleado WHERE empleg = " & l_empleg
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_evaluador = l_rs("ternro")
	end if
	l_rs.close
	
	l_sql = " SELECT evacabnro FROM evacab WHERE evacab.evaevenro = " & l_evaevenro & " AND   evacab.empleado = " & l_empleado
	rsOpen l_rs, cn, l_sql, 0 
	do while not l_rs.eof 
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "UPDATE evadetevldor SET evaluador = " & l_evaluador & " WHERE evadetevldor.evacabnro = " & l_rs("evacabnro") & "	AND evadetevldor.evatevnro= " & l_evatevnro
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs.MoveNext
	loop	
	l_rs.close

	i = i + 2
loop
'cn.CommitTrans

Set l_rs = Nothing
Set cn = Nothing


Response.write "<script>alert('Operación Realizada');window.close();</script>"
%>
