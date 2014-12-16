<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' Modificado: 04-10-2004 CCRossi Agregar campo evaresuEJEM, para ABN...
' Modificado: 29-11-2004 CCRossi Si es ABN, y si el evaluiador es ROL EVALUADOR
' entonces, si evacab.tieneobj=0, grabar puntaje de competencias en evacab.


' variables
' parametros de entrada ----------------------------------------
  Dim l_evafacnro
  Dim l_evldrnro
  Dim l_evaresudesc
  Dim l_evaresuejem
  Dim l_evatrnro
  dim l_evacabnro
  dim l_mostrar
  dim l_evatevnro
  
  dim l_puntaje
  dim l_cantidad
  dim l_tieneobj  

  dim l_decimales
  dim l_entero
  
    		       
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evafacnro	= Request.QueryString("evafacnro")
  l_evaresudesc = request.querystring("evaresudesc")
  l_evaresuejem = request.querystring("evaresuejem")
  l_mostrar		= request.querystring("mostrar")
  
  if len(trim(l_evaresudesc)) <> 0 then
     l_evaresudesc = left(trim(request.querystring("evaresudesc")),200)
   end if 
  if len(trim(l_evaresuejem)) <> 0 then
     l_evaresuejem = left(trim(request.querystring("evaresuejem")),300)
   end if 
  l_evldrnro    = request.querystring("evldrnro")
  l_evatrnro    = request.querystring("evatrnro")

  if l_evatrnro="0" then
	l_evatrnro="null"
  end if

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

'Response.write("l_evaresudesc=")
'Response.write("l_evafacnro=")
'Response.write(l_evafacnro)
'Response.write("<br>")
'Response.write("l_evldrnro=")
'Response.write(l_evldrnro)

'BODY ----------------------------------------------------------


	l_sql = "UPDATE evaresultado SET "
	l_sql = l_sql & " evaresudesc = '"		   & l_evaresudesc & "',"
	l_sql = l_sql & " evaresuejem = '"		   & l_evaresuejem & "',"
	l_sql = l_sql & " evatrnro    =  "		   & l_evatrnro 
	l_sql = l_sql & " WHERE evaresultado.evafacnro = "  & l_evafacnro
	l_sql = l_sql & " AND   evaresultado.evldrnro = "  & l_evldrnro
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	'Response.write("<br>")
	'Response.write(l_sql)

	
	if cejemplo=-1 then ' es ABN!
	
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evacab.evacabnro, evatevnro, tieneobj  "
		l_sql = l_sql & " FROM  evacab "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & "	     AND evadetevldor.evldrnro = "  & l_evldrnro
		 rsOpen l_rs, cn, l_sql, 0
		if not l_rs.EOF then
			l_evacabnro =l_rs("evacabnro")
			l_evatevnro =l_rs("evatevnro")
			l_tieneobj  =l_rs("tieneobj")
		end if
		l_rs.close
		set l_rs=nothing
'		Response.Write("<script>alert('cevaluador "& cevaluador &"');</script>")
'		Response.Write("<script>alert('evatevnro "& l_evatevnro &"');</script>")
		
		'mostrar=1 esignifica que el rol es el LOGEADO
		'si el Evaluador es ROL EVALUADOR (cevaluador)
		'y si no tiene objetivos (tieneobj=0)
		
		if cint(cevaluador)=cint(l_evatevnro) and (l_tieneobj=0) then ' Es evaluador y es el logeado
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT sum(evatrvalor) puntaje, count(evaresu.evatrnro) cantidad "
			l_sql = l_sql & " FROM  evaresu "
			l_sql = l_sql & " INNER JOIN evaresultado ON evaresu.evatrnro = evaresultado.evatrnro "
			l_sql = l_sql & " INNER JOIN evatipresu  ON evatipresu.evatrnro = evaresultado.evatrnro "
			l_sql = l_sql & "	AND evaresultado.evldrnro = "  & l_evldrnro
			 rsOpen l_rs, cn, l_sql, 0
			if not l_rs.EOF then
				l_puntaje  =l_rs("puntaje")
				l_cantidad =l_rs("cantidad")
			end if
			l_rs.close
			set l_rs=nothing
			if l_cantidad <>0 and trim(l_cantidad)<>"" and not isnull(l_cantidad) then
    			l_puntaje = cdbl(l_puntaje) / cdbl(l_cantidad)
    		else	
    			l_puntaje = null
    		end if	
 '   		Response.Write("<script>alert('puntaje "& l_puntaje &"');</script>")
    		if trim(l_puntaje)<>"" and not isnull(l_puntaje) then
    		   l_entero = Cint(l_puntaje)
    		   if l_puntaje < l_entero then
    				if (l_puntaje + 0.5)  > l_entero then
    					l_puntaje = l_entero 
    					else
    					if l_puntaje  < l_entero then
    						l_puntaje = l_entero - 0.5
    					end if	
    				end if
    		   else
	   		   		if l_puntaje  > l_entero then
   						 l_puntaje = l_entero + 0.5
    				end if
    		   end if
    		   
    		   
'    		   Response.Write("<script>alert('entero "& l_entero &"');</script>")
'    		   Response.Write("<script>alert('puntaje "& l_puntaje &"');</script>")
    		   
    		end if
			l_sql = "UPDATE evacab SET "
			l_sql = l_sql & " puntaje = " & PasarComaaPunto(l_puntaje) 
			l_sql = l_sql & " WHERE evacab.tieneobj = 0 "
			l_sql = l_sql & "   AND evacab.evacabnro = " & l_evacabnro
			set l_cm = Server.CreateObject("ADODB.Command")  
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			'Response.Write l_sql
			
			'Response.Write("<script>alert('puntaje "& l_puntaje &"');</script>")
		end if	
		
	end if	
	
	Response.write " <script> parent.Promedio();window.close();</script>"
%>
