<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : licencias_emp_gti_05.asp
Descripcion    : Modulo que se encarga de buscar la cantidad de dias en el anio para un tipo de licencia
Fecha Creacion : 25/03/2004
Autor          : Scarpa D.
Modificacion   :
  29/03/2004 - Scarpa D. - Se agrego un caso para el calculo de licencias por vacaciones
  18/10/2004 - Scarpa D. - Cambio en la forma de calculo de los dias
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_sql
Dim l_rs

dim l_vacnro
Dim l_vacanio
Dim l_corresp
Dim l_opcion
Dim l_canttomados
Dim l_cantturismo

dim l_ternro
dim l_desde
dim l_hasta
Dim l_emp_licnro
dim l_tipo
Dim l_tdnro

Dim aniofin
Dim anioini

Dim aniofin2
Dim anioini2

Dim l_cantidad 
Dim l_tope
Dim l_cantactual

Dim l_tdlimmes
Dim l_tdliman
Dim l_tdcorrido

Dim l_correcto
Dim l_msg

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

dim leg
leg = l_ess_empleg
l_ternro = l_ess_ternro

l_desde  		= request("desde")
l_hasta      	= request("hasta")
l_tdnro  		= request("tdnro")
l_tipo  		= request("tipo")
l_emp_licnro	= request("emplicnro")
l_vacnro        = request("vacnro")
l_cantactual    = request("cantidad")
l_ternro        = request("ternro")

'Busco los datos del tipo de dia

l_sql = "SELECT * FROM tipdia "
l_sql = l_sql & " WHERE tdnro = " & l_tdnro 

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	l_tdlimmes  = l_rs("tdinteger4")
	l_tdliman   = l_rs("tdliman")

	'Si no tiene tope le asigno el maximo
	if isNull(l_tdliman) then
		   l_tdliman = 365	
	else
		if CInt(l_tdliman) = 0 then
		   l_tdliman = 365
		end if
	end if
	
	'Si no tiene tope le asigno el maximo
	if isNull(l_tdlimmes) then
		   l_tdlimmes = 365	
	else
		if CInt(l_tdlimmes) = 0 then
		   l_tdlimmes = 365
		end if
	end if	
	
	l_tdcorrido = l_rs("tdsuma")
end if

l_rs.close

'Si es una licencia de vacaciones busco el periodo al que pertenece
'l_vacnro = ""
'
'if CInt(l_tdnro) = 2 then
'	l_sql = "SELECT * FROM vacacion "
'	l_sql = l_sql & " WHERE vacanio = " & year(CDate(l_desde))
'	
'	rsOpen l_rs, cn, l_sql, 0 
'	
'	if not l_rs.eof then
'	   l_vacnro = l_rs("vacnro")
'	else
'	   l_vacnro = ""	
'	end if
'	
'	l_rs.close
'end if

aniofin = "31/12/" & year(CDate(l_desde))
anioini = "01/01/" & year(CDate(l_desde))

aniofin2 = CDate(aniofin)
anioini2 = CDate(anioini)

'Calculo la cantidad de dias tomados en el anio
if l_vacnro = "" then

		l_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
		l_sql = l_sql & " FROM emp_lic "
		l_sql = l_sql & " WHERE emp_lic.empleado="& l_ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
		l_sql = l_sql & " and tdnro = " & l_tdnro
	    l_sql = l_sql & " AND emp_lic.licestnro= 2 "
		if (l_tipo ="M") then
		l_sql = l_sql & " and emp_licnro <>" & l_emp_licnro
		end if
		
		rsOpen l_rs, cn, l_sql, 0 
		
		l_cantidad = 0
		
		do until l_rs.eof 
		
			if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
			   (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) >= 0) then
			   
			   l_cantidad = l_cantidad + CInt(l_rs("elcantdias"))
			
			else
			   if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) < 0) and _
			      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) >= 0) and _
			      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(anioini2)) <= 0) then
				  
			      l_cantidad = l_cantidad + DateDiff("d",CDate(anioini2),CDate(l_rs("elfechahasta"))) + 1 
				  
			   else
			      if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) <= 0) and _
			         (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) < 0)  and _
			         (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(aniofin2)) >= 0) then
					 
			         l_cantidad = l_cantidad + DateDiff("d",CDate(l_rs("elfechadesde")),CDate(aniofin2)) + 1
					 
				  else
			         if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(anioini2)) > 0) and _
			            (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(aniofin2)) < 0) then
						
			            l_cantidad = l_cantidad + DateDiff("d",CDate(anioini2),CDate(aniofin2)) + 1
						
					 end if
				  end if
			   end if
			end if
		
		   l_rs.moveNext
		loop
		
		l_rs.close
		
'		response.write "//" & l_sql

else

    if l_vacnro <> "0" then

		'Busco los dias correspondientes del empleado
		if l_vacnro <> "" then
			l_sql = "SELECT * FROM vacdiascor "
			l_sql = l_sql & " WHERE vacnro = " & l_vacnro 
			l_sql = l_sql & "   AND ternro = " & l_ternro
			
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
				l_corresp   = l_rs("vdiascorcant")
			else
			    l_corresp	= 0
			end if
			
			l_rs.close
		end if
		
		'Busco la cantidad de dias tomados de la vacaciones
		l_cantidad = 0
		
		if l_vacnro <> "" then
			l_sql =         " SELECT * "
			l_sql = l_sql & " FROM lic_vacacion "
			l_sql = l_sql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
			l_sql = l_sql & " WHERE vacnro = " & l_vacnro
			l_sql = l_sql & " AND emp_lic.empleado= " & l_ternro
		    l_sql = l_sql & " AND emp_lic.licestnro= 2 " 
			
			if (l_tipo ="M") then
			   l_sql = l_sql & " AND emp_lic.emp_licnro <>" & l_emp_licnro
			end if
			
			rsOpen l_rs, cn, l_sql, 0 
			
			l_cantidad = 0
			
			do until l_rs.eof 
			   l_cantidad = l_cantidad + CInt(l_rs("elcantdias"))		
			
			   l_rs.moveNext
			loop
			
			l_rs.close
		end if
		
		l_canttomados = l_cantidad
		
	    l_cantidad = l_canttomados
		l_tdliman  = l_corresp 

	end if

end if

if (CInt(l_cantidad) + CInt(l_cantactual)) > CInt(l_tdliman) then

   l_correcto = false
   l_msg = "La cantidad de dias supera el tope establecido para el tipo de licencia."

else
   if CInt(l_tdlimmes) < CInt(l_cantactual) then

	   l_correcto =  false
	   l_msg = "La cantidad de días supera el máximo por evento."

   else

	   l_correcto =  true
	   l_msg = ""
	   
	   if CInt(l_tdcorrido) <> 0 then
	      if (WeekDay(CDate(l_desde)) = 1) OR (WeekDay(CDate(l_desde)) = 7) then
			 l_correcto =  false
			 l_msg = "La licencia no puede comenzar en sabado ó domingo."
		  end if
	   end if

	end if

end if

if l_correcto then%>
<script>
  parent.guardarLicencia();
</script>   
<%else%>
<script>
  alert('<%= l_msg%>');
</script>   
<%end if%>

