<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' Modificado: 	  - LA - arreglo de calculos
'			  11-07-2007 - LA.- se controlo que el valor evarestot, a parte de nulo, venga vacio (cuando se usan competencias sin ponderacion)

on error goto 0

' variables
' parametros de entrada ----------------------------------------
  Dim l_evafacnro
  Dim l_evldrnro
  Dim l_evaresudesc
  Dim l_evaresuejem

  dim l_evacabnro
  dim l_evatevnro

  Dim l_evatrnro
  Dim l_evarespor
  Dim l_evarestot
  dim l_cantidad
    		       
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evafacnro	= Request.QueryString("evafacnro")
  l_evaresudesc = request.querystring("evaresudesc")
  l_evaresuejem = request.querystring("evaresuejem")
  'l_cantidad	= request.querystring("cantidad")
  
  if len(trim(l_evaresudesc)) <> 0 then
     l_evaresudesc = left(trim(request.querystring("evaresudesc")),200)
   end if 
  if len(trim(l_evaresuejem)) <> 0 then
     l_evaresuejem = left(trim(request.querystring("evaresuejem")),300)
   end if 

  l_evldrnro    = request.querystring("evldrnro")
  l_evatrnro    = request.querystring("evatrnro")
  l_evarespor   = request.querystring("evarespor")
  l_evarestot   = request.querystring("evarestot")
  
  if l_evatrnro="0" then
	l_evatrnro="null"
  end if
  if  trim(l_evarespor)="" then 'l_evarespor="0" or 
	l_evarespor="null"
  end if
 ' if l_evarestot="0" then
	'l_evarestot="null"
 ' end if
'___________________________________________________________________________________
function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	

'Response.write("l_evaresudesc=") 'Response.write("l_evafacnro=")
'Response.write(l_evafacnro) 'Response.write("<br>")
'Response.write("l_evldrnro=") 'Response.write(l_evldrnro)

'___________________________________________________________________________________
'function PasarComaAPunto(valor) 
	'dim l_numero
	'dim l_ubicacion
	'dim l_entero
	'dim l_decimal
	'l_numero = trim(valor)
	'l_ubicacion = InStr(l_numero, ",")
	'if l_ubicacion > 1 then
		'l_ubicacion = l_ubicacion  - 1
		'l_entero = left(l_numero, l_ubicacion)
		'l_ubicacion = l_ubicacion  + 1
		'l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	'l_numero = l_entero & "." & l_decimal
    	'PasarComaAPunto = l_numero
    'else
		'PasarComaAPunto = valor
	'end if
'end function	


'BODY ----------------------------------------------------------

l_sql = "UPDATE evaresultado SET "
l_sql = l_sql & " evaresudesc = '"		   & l_evaresudesc & "',"
l_sql = l_sql & " evaresuejem = '"		   & l_evaresuejem & "',"
l_sql = l_sql & " evatrnro    =  "		   & l_evatrnro  
if not isnull(l_evarespor) then
l_sql = l_sql & ", evarespor    = "		   & PasarComaaPunto(l_evarespor)
'else
'l_sql = l_sql & " evarespor    =  "		   & l_evarespor & ","
end if
if not isnull(l_evarestot) and l_evarestot<>"" then
l_sql = l_sql & ", evarestot    =  "		   & PasarComaaPunto(l_evarestot)
'else
'l_sql = l_sql & " evarestot    =  "		   & l_evarestot 
end if
l_sql = l_sql & " WHERE evaresultado.evafacnro = "  & l_evafacnro
l_sql = l_sql & " AND   evaresultado.evldrnro = "  & l_evldrnro
set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
'Response.write("<br>")
'Response.write(l_sql)
	'		l_sql = "UPDATE evacab SET "
	'		l_sql = l_sql & " puntaje = " & PasarComaaPunto(l_puntaje) 
	'		l_sql = l_sql & " WHERE evacab.tieneobj = 0 "
	'		l_sql = l_sql & "   AND evacab.evacabnro = " & l_evacabnro
	'		set l_cm = Server.CreateObject("ADODB.Command")  
	'		l_cm.activeconnection = Cn
	'		l_cm.CommandText = l_sql
	'		cmExecute l_cm, l_sql, 0
	'		'Response.Write l_sql
	'		
	'		'Response.Write("<script>alert('puntaje "& l_puntaje &"');</script>")
	
' Response.write " <script> parent.Promedio("&l_cantidad&");</script>"
Response.write " <script> parent.Promedio();</script>"
%>
