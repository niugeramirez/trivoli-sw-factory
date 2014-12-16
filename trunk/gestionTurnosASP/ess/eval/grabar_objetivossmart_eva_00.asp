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
'Modificacion  : 04-08-2004 - CCRossi - agregar deletes de tablas relacionadas al objetivo.
'Modificacion  : 17-08-2004 - CCRossi - solo grabar el objetivo y la ponderacion
'				 el resto se graba en el coach.
'Modificacion  : 02-11-2004 CCRossi- grabarel objetivo y actualizar estadoseccion
'Modificacion  : 24-11-2004 CCRossi- arreglar error en grabar
'=====================================================================================

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjnro
  Dim l_evaobjdext
  Dim l_evaobjformed
  Dim l_evaobjpond
  Dim l_evaperfijo
  Dim l_evasgtotexto  
  Dim l_evatipobjnro
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_sqlb ' para hacer las bajas
  Dim l_rs
  Dim l_rs1
  Dim l_tipo  
  
' locales
  Dim l_evacabnro 
  Dim l_evatevnro
  Dim l_evaluador
  dim l_originalevldrnro

  dim l_suma
  dim l_cartel
  
' parametros de entrada
  l_evaobjdext   = left(trim(request.querystring("evaobjdext")),300)
  l_evaobjpond   = request.querystring("evaobjpond")
  l_evldrnro	 = request.querystring("evldrnro")
  l_originalevldrnro = request.querystring("evldrnro")
  l_tipo		 = request.querystring("tipo")
  l_evaobjnro	 = request.querystring("evaobjnro")
  l_evaobjformed	 = request.querystring("evaobjformed")
  
  if trim(l_evaobjformed)<>"" then
  l_evaobjformed	 = left(trim(l_evaobjformed),300)
  end if
  l_evaperfijo	 = request.querystring("evapernro")
  l_evatipobjnro  = request.querystring("evatipobjnro")
  if trim(l_evatipobjnro)="" or l_evatipobjnro=0 then
	l_evatipobjnro = null
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
Case "A":
		' Primero sumar lo que hay para ver si excede el 100%
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaobjetivo.evaobjpond "
		l_sql = l_sql & " FROM evaobjetivo "
		l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
		l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
		l_sql = l_sql & " WHERE evatipoobj.evatipobjnro = " & l_evatipobjnro
		l_sql = l_sql & " AND   evaluaobj.evldrnro = " & l_evldrnro
		'Response.Write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		l_suma=0
		do while not l_rs.eof
			if trim(l_rs("evaobjpond"))<>"" then
			l_suma= l_suma + cdbl(l_rs("evaobjpond"))
			end if
			l_rs.MoveNext
		loop
		l_rs.close
		set l_rs=nothing
		if (l_suma + l_evaobjpond) > 100 then
			l_evaobjpond = 0
			l_cartel = " Modifique la Ponderación del nuevo Compromiso.\n La sumatoria no puede superar 100%."
		end if
		
		l_sql= "insert into evaobjetivo (evaobjdext,evaobjformed,evaobjpond,evatipobjnro) "
		l_sql = l_sql & " values ('" & trim(l_evaobjdext) & "','" & trim(l_evaobjformed) & "'," & l_evaobjpond & "," & l_evatipobjnro &")"
case "M":
		l_sql = "UPDATE evaobjetivo SET "
		l_sql = l_sql & " evaobjdext     = '" & trim(l_evaobjdext) & "',"
		l_sql = l_sql & " evaobjformed      = '" & trim(l_evaobjformed)  & "',"
		l_sql = l_sql & " evaobjpond     = " & l_evaobjpond  & " "
		l_sql = l_sql & " WHERE evaobjetivo.evaobjnro = "  & l_evaobjnro
case "B":
		l_sql = "DELETE from evaobjetivo where evaobjetivo.evaobjnro = "  & l_evaobjnro
end select

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Response.Write l_sql

if trim(l_evaobjnro)="" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("evaobjnro","evaobjetivo")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_evaobjnro=l_rs("evaobjnro")
	end if	
	l_rs.Close
	Set l_rs = Nothing
end if

'Crear registro de la relacion ================================================
select case l_tipo
Case "A":
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaluaobj "
	l_sql = l_sql & " WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		l_sql= "insert into evaobjsgto (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
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
Case "A","M":
'crear los evaluaobj para el mismo Evaluador y disntos evldnro (para coach)
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

if cmismo_objetivo=1 then ' todos los evaluadores evaluan los mismo objetivos para el mismo EVALUADO
'crear los evaluaobj para LOS OTROS Evaluadores 
'busco los evldrnro de los otros evatevnro
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evldrnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evatevnro  <> " & l_evatevnro
l_sql = l_sql & " AND   tipsecobj=-1" 
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	l_evldrnro= l_rs("evldrnro")
	'busco el objetivo si ya existe para este evldrnro
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaluaobj "
	l_sql = l_sql & " WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
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
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaobjsgto (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjnro &")"
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
end if


end select
'response.write "<script> parent.parent.document.estado.location='estadoseccion_eva_ag_00.asp?evldrnro="& l_originalevldrnro& "'&logeado=-1';</script>"
'response.write "<script> alert('xxx'); </script>"
response.write "<script> //parent.parent.document.estado.location.reload();</script>"
if trim(l_cartel)<>"" then
response.write "<script> alert('"&l_cartel&"'); </script>"
end if
response.write "<script> parent.location.reload(); </script>"
%>
