<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'======================================================================================
'Archivo		: grabar_competencias_evaluacion_00.asp
'Descripción	: Graba los resultados de la evaluac de las competencias
'Autor			: LAmadio.
'Fecha			: 08-03-2005
'Modificado		:
'======================================================================================
'on error goto 0

' variables
' parametros de entrada ----------------------------------------
  Dim l_evafacnro
  Dim l_evldrnro
  Dim l_evaresudesc
  'Dim l_evaresuejem
  Dim l_evatrnro
  dim l_evacabnro
  dim l_mostrar
  dim l_evatevnro
  dim l_campo
  
    		       
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evafacnro	= Request.QueryString("evafacnro")
  l_evaresudesc = request.querystring("evaresudesc")
  'l_evaresuejem = request.querystring("evaresuejem")
  l_mostrar		= request.querystring("mostrar")
  
  l_campo = request.querystring("campo")
  
  'l_evaresuejem=""
  
  if len(trim(l_evaresudesc)) <> 0 then
     l_evaresudesc = left(trim(request.querystring("evaresudesc")),200)
   end if 
  'if len(trim(l_evaresuejem)) <> 0 then
   '  l_evaresuejem = left(trim(request.querystring("evaresuejem")),300)
   'end if 
  l_evldrnro    = request.querystring("evldrnro")
  'l_evatrnro    = request.querystring("evatrnro")

  'if l_evatrnro="0" then
	'l_evatrnro="null"
  'end if


'Response.write("l_evaresudesc=")
'Response.write("l_evafacnro=")
'Response.write(l_evafacnro)
'Response.write("<br>")
'Response.write("l_evldrnro=")
'Response.write(l_evldrnro)

'BODY ----------------------------------------------------------

	l_sql = "UPDATE evaresultado SET "
	l_sql = l_sql & " evaresudesc = '"  & l_evaresudesc & "'"
	'l_sql = l_sql & " evaresuejem = '"  & l_evaresuejem & "',"
	'l_sql = l_sql & " evatrnro    =  "  & l_evatrnro 
	l_sql = l_sql & " WHERE evaresultado.evafacnro = "  & l_evafacnro
	l_sql = l_sql & " AND   evaresultado.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	response.write "<script> parent.document.datos."&l_campo&".focus();</script>"
	'response.write " <script> alert('"& l_campo &"')</script>"
	Response.write " <script> window.close();</script>"
%>
