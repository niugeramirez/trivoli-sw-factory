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
  Dim l_evaobjdext
  Dim l_evaobjformed
  Dim l_evaobjpond
  Dim l_evaperfijo
  Dim l_evaobjalcanza
  Dim l_evasgtotexto  

' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_sqlb ' para hacer las bajas
  Dim l_rs
  Dim l_rs1
  Dim l_tipo  ' solo viene con B o M
  
' locales
  Dim l_evacabnro 
  Dim l_evatevnro
  Dim l_evaluador

  
' parametros de entrada
  l_evaobjdext   = left(trim(request.querystring("evaobjdext")),300)
  l_evaobjpond   = request.querystring("evaobjpond")
  l_evaobjalcanza= request.querystring("evaobjalcanza")
  l_evldrnro	 = request.querystring("evldrnro")
  l_tipo		 = request.querystring("tipo")
  l_evaobjnro	 = request.querystring("evaobjnro")
  l_evaperfijo	 = request.querystring("evapernro")
  
  l_evasgtotexto  = request.querystring("evasgtotexto")
  if trim(l_evasgtotexto)<>"" then
  l_evaobjformed  = left(trim("evaobjformed"),200)
  end if
  
  
'BODY ----------------------------------------------------------
if l_tipo="B" then
		l_sqlb = "DELETE FROM evaluaobj  "
		l_sqlb = l_sqlb & " where evaluaobj.evaobjnro= " & l_evaobjnro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sqlb
		cmExecute l_cm, l_sqlb, 0

		l_sqlb = "DELETE FROM evaplan  "
		l_sqlb = l_sqlb & "     WHERE evaplan.evaobjnro = " & l_evaobjnro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sqlb
		cmExecute l_cm, l_sqlb, 0
 
		l_sqlb = "DELETE FROM evaobjsgto  "
		l_sqlb = l_sqlb & " WHERE evaobjsgto.evaobjnro =" & l_evaobjnro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sqlb
		cmExecute l_cm, l_sqlb, 0
end if

select case l_tipo
case "M":
		l_sql = "UPDATE evaobjetivo SET "
		l_sql = l_sql & " evaobjdext     = '" & trim(l_evaobjdext) & "',"
		l_sql = l_sql & " evaobjformed   = '" & trim(l_evaobjformed) & "',"		
		l_sql = l_sql & " evaobjpond     = " & l_evaobjpond  
		l_sql = l_sql & " WHERE evaobjetivo.evaobjnro = "  & l_evaobjnro
case "B":
		l_sql = "DELETE from evaobjetivo where evaobjetivo.evaobjnro = "  & l_evaobjnro
end select

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

'Crear registro de la relacion ================================================
select case l_tipo
Case "M":
	l_sql = "UPDATE evaluaobj SET "
	l_sql = l_sql & " evaobjalcanza     = " & l_evaobjalcanza
	l_sql = l_sql & " WHERE evaobjnro = "  & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_sql = "UPDATE evaobjsgto SET "
	l_sql = l_sql & " evasgtotexto     = '" & trim(l_evasgtotexto) & "'"
	l_sql = l_sql & " WHERE evaobjsgto.evaobjnro = "  & l_evaobjnro
	l_sql = l_sql & " AND   evaobjsgto.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
end select

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = "& l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evacabnro =l_rs("evacabnro")
	l_evatevnro =l_rs("evatevnro")
	l_evaluador =l_rs("evaluador")
end if	
l_rs.Close
Set l_rs = Nothing

select case l_tipo
Case "M":
' crear los evaluaobj y evaobjsgto para el resto de los evldrnro, 
'    de la cabecera, evaluador y objetivo
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evldrnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
l_sql = l_sql & " AND   evaluador = " & l_evaluador
l_sql = l_sql & " AND   evldrnro  <> " & l_evldrnro
l_sql = l_sql & " AND   tipsecobj=-1" 
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaluaobj "
	l_sql = l_sql & " WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_rs("evldrnro")
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_rs("evldrnro") & "," & l_evaobjnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_rs1.Close
		set l_rs1=nothing
	end if
	
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaobjsgto "
	l_sql = l_sql & " WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_rs("evldrnro")
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaobjsgto (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_rs("evldrnro") & "," & l_evaobjnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_rs1.Close
		set l_rs1=nothing
	end if
	l_rs.MoveNext
loop
l_rs.Close
Set l_rs = Nothing
end select

response.write "<script> parent.location.reload(); </script>"
%>
