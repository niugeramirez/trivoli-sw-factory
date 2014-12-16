<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include file="vacaciones_calculo_gti.asp"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : licencias_emp_gti_06.asp
Descripcion    : Modulo que se contar dias
Fecha Creacion : 29/03/2004
Autor          : Scarpa D.
Modificacion   :
  06/05/2004 - Scarpa D. - No sumar los fines de semanas
  18/10/2004 - Scarpa D. - Cambio en la forma de calculo de los dias
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

dim l_rs
dim l_rs2
dim m_rs
dim l_sql

dim l_desde
dim l_hasta
dim l_cant
dim l_tipo
dim l_salida
dim l_actual
dim l_i
Dim l_tdcorrido
Dim l_tdnro
dim leg
dim l_ternro
dim l_pais
Dim l_vacnro
Dim l_tipoLic
Dim l_tipvacnro

Dim lngAlcanGrupo
    lngAlcanGrupo = 2

l_desde  		= request("desde")
l_hasta      	= request("hasta")
l_cant  		= request("cant")
l_tipo  		= request("tipo")
l_tdnro 		= request("tdnro")
l_tipoLic  		= request("tipolic")
l_vacnro  		= request("vacnro")
l_ternro  		= request("ternro")

if l_tipoLic = "" then
   l_tipoLic = "0"
end if

if l_vacnro = "" then
   l_vacnro = "0"
end if

if l_ternro = "" then
   l_ternro = "0"
end if

Set l_rs   = Server.CreateObject("ADODB.RecordSet")
Set l_rs2  = Server.CreateObject("ADODB.RecordSet")

'---------------------------------------------------------------------------------------------
'Busco los datos del empleado
leg = l_ess_empleg
l_ternro = l_ess_ternro

'---------------------------------------------------------------------------------------------
'Busco los datos del tipo de dia

l_sql = "SELECT * FROM tipdia "
l_sql = l_sql & " WHERE tdnro = " & l_tdnro 

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	l_tdcorrido = l_rs("tdsuma")
end if

l_rs.close

if CInt(l_tdnro) <> 2 then

	'---------------------------------------------------------------------------------------------
	'Calculo la cantidad de dias
	if l_tipo = "SUMAR" then
	   l_actual = CDate(l_desde)
	   l_i = 0
	   
	   while l_i < CInt(l_cant)
	      if CInt(l_tdcorrido) = 0 then
		      if weekday(l_actual) <> 1 AND weekday(l_actual) <> 7 then
			     if not esFeriado(l_actual, l_pais) then
			        l_salida =  l_actual
		   	        l_i = l_i + 1
				 end if
			  end if
		  else
			  l_salida =  l_actual
		   	  l_i = l_i + 1
		  end if
		  l_actual = DateAdd("d", 1,CDate(l_actual))
	   wend 
	
	else
	
	   l_actual = CDate(l_desde)
	   l_i = 0
	
	   do
	      if CInt(l_tdcorrido) = 0 then
		      if weekday(l_actual) <> 1 AND weekday(l_actual) <> 7 then
			     if not esFeriado(l_actual, l_pais) then
		   	        l_i = l_i + 1
				 end if
			  end if
		  else
		   	  l_i = l_i + 1
		  end if
	
		  l_actual = DateAdd("d", 1,CDate(l_actual))
	   loop until DateDiff("d",CDate(l_actual), CDate(l_hasta) ) < 0	  
	
	   l_salida = l_i
	
	end if
	
else

    'Busco el tipo de vacacion
	l_sql = " SELECT tipvacnro FROM vacdiascor "
	l_sql = l_sql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
	l_sql = l_sql & " WHERE vacdiascor.ternro = " & l_ternro
	l_sql = l_sql & "   AND vacdiascor.vacnro = " & l_vacnro
	
	rsOpen l_rs, cn, l_sql, 0
	if l_rs.eof then
	   l_tipvacnro = 0
	else
	   l_tipvacnro = l_rs("tipvacnro")
	end if
	l_rs.close

    'Inicializo las variables del modulo
    call iniTipoVac(l_tipvacnro)
	
	Dim l_total
	Dim l_totalFer	

	if l_tipo = "SUMAR" then
	
	  'Suma dias a una fecha
	  call busqFecha(l_desde,l_cant,l_hasta,l_total, l_totalFer)  
	  l_salida = l_hasta
	
	else
	
	  'Cantidad de dias entre dos fecha
	  call cantDias(l_desde,l_hasta,l_cant,l_total,l_totalFer)
	  l_salida = l_cant
	  
	end if  

end if

if l_tipo = "SUMAR" then
%>
<script>
  parent.document.datos.elfechahasta.value = '<%= l_salida%>';
</script>   
<%else%>
<script>
  parent.document.datos.elcantdias.value = '<%= l_salida%>';
</script>   
<%end if%>



