<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : grabar_objetivossmart_eva_00.asp
'Objetivo : grabar de objetivos smart de evaluacion
'Fecha	  : 14-05-2004
'Autor	  : CCRossi
'Modificacion  : 04-08-2004 - CCRossi - agrgegar deletes de tablas relacionadfas al objetivo.
'=====================================================================================

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjnro
  Dim l_evaobjalcanza
  Dim l_evasgtotexto  

' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql

  dim l_rs1
  
  
' parametros de entrada
  l_evaobjalcanza = request.querystring("evaobjalcanza")
  l_evldrnro	  = request.querystring("evldrnro")
  l_evaobjnro	  = request.querystring("evaobjnro")
  l_evasgtotexto  = request.querystring("evasgtotexto")
  
  if trim(l_evasgtotexto)<>"" then
	l_evasgtotexto=left(l_evasgtotexto,200)
  end if	
'BODY ----------------------------------------------------------

'actualizar NOTA en registro de la relacion ================================================
l_sql = "UPDATE evaluaobj SET "
l_sql = l_sql & " evaobjalcanza     = " & l_evaobjalcanza
l_sql = l_sql & " WHERE evaobjnro = "  & l_evaobjnro
l_sql = l_sql & " AND   evldrnro  = "  & l_evldrnro
set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
	
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & " FROM evaobjsgto "
l_sql = l_sql & " WHERE evaobjsgto.evaobjnro = "  & l_evaobjnro
l_sql = l_sql & " AND   evaobjsgto.evldrnro = "  & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0 
if l_rs1.EOF then
	l_sql = "INSERT INTO evaobjsgto (evaobjnro,evldrnro,evasgtotexto) "
	l_sql = l_sql & " VALUES (" 
	l_sql = l_sql & l_evaobjnro & ","
	l_sql = l_sql & l_evldrnro  & ",'"
	l_sql = l_sql & trim(l_evasgtotexto) & "')"
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
else
	l_sql = "UPDATE evaobjsgto SET "
	l_sql = l_sql & " evasgtotexto     = '" & trim(l_evasgtotexto) & "'"
	l_sql = l_sql & " WHERE evaobjsgto.evaobjnro = "  & l_evaobjnro
	l_sql = l_sql & " AND   evaobjsgto.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end if	
l_rs1.Close
set l_rs1=nothing

cn.close
set cn=nothing

'response.write "<script> parent.location.reload(); </script>"
%>
