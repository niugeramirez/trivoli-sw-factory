<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<% 
'=====================================================================================
'Archivo  : grabar_planobjetivossmart_eva_00.asp
'Objetivo : grabar plan de objetivos smart de evaluacion
'Fecha	  : 18-06-2004
'Autor	  : CCRossi
'=====================================================================================

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_evaobjnro
  Dim l_evaplnro
  Dim l_aspectomejorar
  Dim l_planaccion
  Dim l_planfecharev
  Dim l_recursos
  Dim l_ayuda  

' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_tipo  
  Dim l_rs

' locales
  Dim l_evaseccnro
  Dim l_evacabnro
    
' parametros de entrada
  l_aspectomejorar  = left(trim(request.querystring("aspectomejorar")),200)
  l_planaccion		= left(trim(request.querystring("planaccion")),200)
  l_planfecharev	= request.querystring("planfecharev")
  l_recursos		= left(trim(request.querystring("recursos")),200)
  l_ayuda			= left(trim(request.querystring("ayuda")),200)
  l_evldrnro		= request.querystring("evldrnro")
  l_evaobjnro		= request.querystring("evaobjnro")
  l_evaplnro		= request.querystring("evaplnro")
  l_tipo			= request.querystring("tipo")
  l_evaseccnro		= request.querystring("evaseccnro")
  
if  trim(l_planfecharev)<>"" then
	l_planfecharev = cambiafecha(l_planfecharev,"","")
else
	l_planfecharev	="null"
end if	

'Buscar la seccion
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaseccnro , evacabnro FROM evadetevldor WHERE evldrnro = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evaseccnro = l_rs("evaseccnro")
	l_evacabnro  = l_rs("evacabnro")
end if
l_rs.close
set l_rs=nothing

'BODY ----------------------------------------------------------
select case l_tipo
Case "A":
		l_sql= "insert into evaplan (evaobjnro,evldrnro,aspectomejorar,planaccion,planfecharev,recursos,ayuda) "
		l_sql = l_sql & " values (" & l_evaobjnro &","& l_evldrnro & ",'" & trim(l_aspectomejorar) & "','" & trim(l_planaccion) & "'," & l_planfecharev & ",'"&l_recursos&"','" & l_ayuda &"')"
case "M":
		l_sql = "UPDATE evaplan SET "
		l_sql = l_sql & " aspectomejorar ='" & l_aspectomejorar & "',"
		l_sql = l_sql & " planaccion     ='" & l_planaccion		& "',"
		l_sql = l_sql & " planfecharev   = "  & l_planfecharev  & " ,"
		l_sql = l_sql & " recursos		 ='"  & l_recursos		& "', "
		l_sql = l_sql & " ayuda			 ='"  & l_ayuda			& "' "
		l_sql = l_sql & " WHERE evaplan.evaplnro = "  & l_evaplnro
case "B":
		l_sql = "UPDATE evaplan SET "
		l_sql = l_sql & " aspectomejorar ='',"
		l_sql = l_sql & " planaccion     ='',"
		l_sql = l_sql & " planfecharev   =null,"
		l_sql = l_sql & " recursos		 ='', "
		l_sql = l_sql & " ayuda			 ='' "
		l_sql = l_sql & " WHERE evaplan.evaplnro = "  & l_evaplnro
end select

set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

if ccodelco=-1 and trim(l_evaobjnro)<>"" then
	'crear o actualizar los mismo textos para el resto de los actores
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evadetevldor.evldrnro, evaplan.evaobjnro "
	l_sql = l_sql & " FROM evadetevldor "
	l_sql = l_sql & " INNER JOIN evaplan ON evaplan.evldrnro=evadetevldor.evldrnro "
	l_sql = l_sql & " WHERE evaobjnro = " & l_evaobjnro
	l_sql = l_sql & " AND   evadetevldor.evldrnro <> " & l_evldrnro
	l_sql = l_sql & " AND   evadetevldor.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " AND   evadetevldor.evacabnro  =" & l_evacabnro
	'Response.Write l_sql
	rsOpen l_rs, cn, l_sql, 0
	do while not l_rs.eof 
		
		l_sql = "UPDATE evaplan SET "
		l_sql = l_sql & " aspectomejorar ='" & l_aspectomejorar & "',"
		l_sql = l_sql & " planfecharev   = "  & l_planfecharev  & " "
		l_sql = l_sql & " WHERE evaplan.evldrnro  = "  & l_rs("evldrnro")
		l_sql = l_sql & " AND   evaplan.evaobjnro = "  & l_evaobjnro
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		l_rs.MoveNext
	loop
	l_rs.Close
	set l_rs=nothing	
end if

cn.Close
set cn = Nothing

if l_tipo="B" then
response.write "<script> parent.location.reload(); </script>"
end if
%>
