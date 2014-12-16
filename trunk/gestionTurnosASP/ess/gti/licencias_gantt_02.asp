<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/util.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo        : licencias_gantt_01.asp
Descripción    : Control Licencias Empleados - salida HTML
Autor          : Scarpa D.
Fecha Creacion : 30/08/2004
Modificado     : 
-->

<% 
const Color_Feriado   = "#ccffff"
const Color_FinSemana = "#e0e0e0"
const Color_Aprobado  = "#ccffcc"
const Color_Pendiente = "#edbf6b"

on error goto 0

Const l_Max_Lineas_X_Pag = 10000
const l_nro_col = 4

' Variables
Dim l_ternro
Dim leg
dim l_rs
dim l_sql
dim l_sqlTer

Set l_rs  = Server.CreateObject("ADODB.RecordSet")

l_ternro = l_ess_ternro 
leg = l_ess_empleg

Dim l_arrTmp 
Dim l_arrTmp2
Dim l_tituloR 
Dim l_poner_titulo
Dim l_pagina
Dim l_descripcion
Dim l_factual
Dim l_hay_datos
Dim l_emp_actual
Dim l_salir
Dim l_i
Dim l_empleado
Dim l_mesant
Dim l_vacnro

Dim l_corresp1
Dim l_cantidad1
Dim l_total1
Dim l_corresp2
Dim l_cantidad2
Dim l_total2

Dim l_rs2
Dim l_rs20

Dim l_cm

Dim l_sqlemp

Dim l_nrolinea
Dim l_nropagina
Dim l_totalemp

Dim l_encabezado
Dim l_corte
Dim l_cambioEmp
Dim l_conc_detdom
Dim l_listaproc
Dim l_procant
Dim l_terant
Dim l_pliqdesc
Dim l_pliqant
Dim l_pliqmesant
Dim l_pliqanioant

'Parametros
 Dim l_filtro ' Viene el filtro comun: empest, legajo, 
 Dim l_orden
 Dim l_desde	
 Dim l_hasta	
 Dim l_cant_dias
 Dim l_anio
 
l_filtro = request("filtro")
l_orden  = request("orden")
l_anio   = request("anio")

if l_orden = "" then
   l_orden = "terape"
end if

if l_anio = "" then
   l_anio = year(date)
end if

l_desde = "01/01/" & l_anio
l_hasta = "31/12/" & l_anio

l_cant_dias = DateDiff("d",CDate(l_desde),CDate(l_hasta)) + 1

Dim l_arrDias(400)
Dim l_arrEstados(400)
Dim l_arrFeriados(400)
Dim l_arrLicencias(400)
Dim l_cantFeriados
Dim l_tipoDia(3000)
Dim l_tipoDiaDesc(3000)

'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
function diaSemana(fecha)

   if CInt(Day(CDate(fecha))) < 10 then
      diaSemana = "0" & Day(CDate(fecha))   
   else
      diaSemana = Day(CDate(fecha))
   end if

end function 'l_salida(fecha)

'---------------------------------------------------------------------------------------------------------
' FUNCION: esFeriado - calcula si una fecha es feriado
'---------------------------------------------------------------------------------------------------------
function esFeriado(dia,pais)

  Dim l_salida
  Dim l_p
  
  l_salida = false
  
  for l_p = 0 to (l_cantFeriados - 1)
     if DateDiff("d",CDAte(l_arrFeriados(l_p)),CDate(dia)) = 0 then
	    l_salida = true
	    exit for
	 end if
  next

  esFeriado = l_salida
end function 'esFeriado(dia)

'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
function colorDia(fecha)

   if esFeriado(fecha,170) then
      colorDia = Color_feriado
   else
     if weekday(fecha) = 1 OR weekday(fecha) = 7 then
        colorDia = Color_finSemana
	 else
        colorDia = ""
	 end if
   end if

end function 'l_salida(fecha)

'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
function descMes(mes)
   Dim l_salida

   select case mes
      case 1: l_salida = "Enero"
      case 2: l_salida = "Febrero"
      case 3: l_salida = "Marzo"
      case 4: l_salida = "Abril"
      case 5: l_salida = "Mayo"
      case 6: l_salida = "Junio"
      case 7: l_salida = "Julio"
      case 8: l_salida = "Agosto"
      case 9: l_salida = "Septiembre"
      case 10: l_salida = "Octubre"
      case 11: l_salida = "Noviembre"
      case 12: l_salida = "Diciembre"
   end select
   
   descMes = l_salida

end function 'descMes(mes)


'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
function obtColor(estado)
   select case CInt(estado)
      case 1
	    obtColor = Color_pendiente
      case 2
	    obtColor = Color_aprobado 
   end select

end function 'obtColor(estado)

'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
sub inicializarDias()

 Dim l_i
 
 for l_i = 1 to l_cant_dias
     l_arrDias(l_i) = 0
     l_arrEstados(l_i) = 0 
	 l_arrLicencias(l_i) = 0
 next

end sub 'inicializarDias()

'----------------------------------------------------------------------------------------------------
'Descripcion:
'----------------------------------------------------------------------------------------------------
sub cargarRango(desde,hasta,tdnro,estado,licencia)

 Dim l_i
 Dim l_modificable
 
 l_factual = CDate(l_desde)
 l_i = 1
 
 if estado = 2 then
    l_modificable = (DateDiff("d",Date(),desde) >= 0)
 else
    l_modificable = true
 end if
 
 while DateDiff("d",l_factual,CDate(hasta)) >= 0
 
     if (DateDiff("d",CDate(desde),l_factual) >= 0) AND (DateDiff("d",l_factual,CDate(hasta)) >= 0) then
	    l_arrDias(l_i) = tdnro
		l_arrEstados(l_i) = estado 
		if l_modificable then
	       l_arrLicencias(l_i) = licencia
		else
	       l_arrLicencias(l_i) = 0
		end if
	 end if

	 l_i = l_i + 1
     l_factual = DateAdd("d",1,l_factual)
 wend

end sub 'cargarRango(desde,hasta,tdnro)

'------------------------------------------------------------------------------------------------
' SUB: buscarDatosVacaciones()
'------------------------------------------------------------------------------------------------
sub buscarDatosVacaciones(vacnro,ternro,anio, byRef corresp, byRef cantidad)

	'Busco los dias correspondientes del empleado
	if vacnro <> "" then
		l_sql = "SELECT SUM(vdiascorcant) AS suma FROM vacdiascor "
		l_sql = l_sql & " WHERE ternro = " & ternro
		
		rsOpen l_rs2, cn, l_sql, 0 
		
		corresp	= 0
		if not l_rs2.eof then
		   if not isNull(l_rs2("suma")) then
			  corresp = l_rs2("suma")
		   end if
		end if
	
		l_rs2.close
	
		l_sql =         " SELECT sum(elcantdias) AS total "
		l_sql = l_sql & " FROM emp_lic  "
		l_sql = l_sql & " WHERE tdnro = 2 "
		l_sql = l_sql & " AND emp_lic.elfechadesde < 1/1/" & anio
		l_sql = l_sql & " AND emp_lic.empleado= " & ternro
		l_sql = l_sql & " AND licestnro IN (1,2) "
	
		rsOpen l_rs2, cn, l_sql, 0 
	
		if not l_rs2.eof then
		   if not isNull(l_rs2("total")) then
		      corresp   = corresp -  CInt(l_rs2("total"))
		   end if
		end if
		
		l_rs2.close
	else
        corresp  = 0
	end if
	
	'Busco la cantidad de dias tomados de la vacaciones
	cantidad = 0
	
	if vacnro <> "" then
	
		l_sql =         " SELECT sum(elcantdias) AS total"
		l_sql = l_sql & " FROM emp_lic  "
		l_sql = l_sql & " WHERE tdnro = 2 "
		l_sql = l_sql & " AND emp_lic.empleado= " & ternro
		l_sql = l_sql & " AND licestnro IN (1,2) "	
		l_sql = l_sql & " AND emp_lic.elfechadesde >= 1/1/" & anio
		
		rsOpen l_rs2, cn, l_sql, 0 
		
		cantidad = 0
		
		if not l_rs2.eof then
		   if not isNull(l_rs2("total")) then
		      cantidad = CInt(l_rs2("total"))
		   end if
		end if
		
		l_rs2.close
	else
        cantidad = 0
	end if

end sub 'buscarDatosVacaciones()

'------------------------------------------------------------------------------------------------
' SUB: buscarDatosTurismo()
'------------------------------------------------------------------------------------------------
sub buscarDatosTurismo(vacnro,ternro,anio, byRef corresp, byRef cantidad)

    if vacnro <> "" then
        Dim aniofin
		Dim anioini
		
	    aniofin = "31/12/" & anio
	    anioini = "01/01/" & anio
	
		l_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
		l_sql = l_sql & " FROM emp_lic "
		l_sql = l_sql & " WHERE emp_lic.empleado="& ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
		l_sql = l_sql & " and licestnro IN (1,2) "
		l_sql = l_sql & " and tdnro = 28 " 
		
		rsOpen l_rs2, cn, l_sql, 0 
		
		'Si no tiene 'licencias de turismo de trabajo' asignadas no puede tomar 'licencias de turismo'
		if l_rs2.eof then
		   corresp = 0
		else
		   corresp = 5
		end if
	
		l_rs2.close
	
		l_sql = "SELECT emp_licnro,elfechadesde,elfechahasta, elcantdias "
		l_sql = l_sql & " FROM emp_lic "
		l_sql = l_sql & " WHERE emp_lic.empleado="& ternro &" and ((elfechadesde >=" & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) & ") "
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta <= " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechahasta >= " & cambiafecha(anioini,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) 
		l_sql = l_sql & " and elfechadesde <= " & cambiafecha(aniofin,"YMD",true) & ") "	
		l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(anioini,"YMD",true)
		l_sql = l_sql & " and elfechahasta >  " & cambiafecha(aniofin,"YMD",true) & ")) "
		l_sql = l_sql & " and licestnro IN (1,2) "
		l_sql = l_sql & " and tdnro = 18 " 
		
		rsOpen l_rs2, cn, l_sql, 0 
		
		'Si no tiene 'licencias de turismo de trabajo' asignadas no puede tomar 'licencias de turismo'
		if l_rs2.eof then
		   cantidad = 0
		else
		   cantidad = 5
		end if
	
		l_rs2.close
	else
	   cantidad = 0
	   corresp  = 0
	end if
	
end sub 'buscarDatosTurismo()

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>

<style>
BODY{
	border: none;
}
TABLE
{
	border : thick solid 1;
}
.TABLE2
{
	border : none solid 0;
}

TH
{
/*	background-color: #333399;*/
/*	COLOR: #ffffff;*/
	FONT-FAMILY: "Arial";
	FONT-SIZE: 9pt;
	FONT-WEIGHT: bold;
	padding : 2 2 2 5;
	width : auto;
}
TR
{
	COLOR: black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: #FFFFFF;
	padding : 2;
	padding-left : 5;
}

.tr
{
	COLOR: black;
	FONT-FAMILY: Verdana;
	FONT-SIZE: 08pt;
	BACKGROUND-COLOR: #DFFFFF;
	padding : 2;
	padding-left : 5;
}


</style>

<script>

function cambioEstado(licencia,estado){
	abrirVentana('sup_licencias_02.asp?origen=control&cabnro='+licencia+'&estado='+estado,'',360,155,',scrollbars=no')
}

function mover(){
  parent.ifrm1.document.body.scrollTop = document.body.scrollTop;
}

</script>

</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onScroll="mover();">
<%

'Busco el periodo de vacaciones

l_sql = "SELECT * FROM vacacion "
l_sql = l_sql & " WHERE vacanio = " & l_anio

rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
   l_vacnro = l_rs("vacnro")
else
   l_rs.close
   
   l_sql = "SELECT vacnro, max(vacanio) FROM vacacion "
   l_sql = l_sql  & " GROUP BY vacnro " 			   
   
   rsOpen l_rs, cn, l_sql, 0 

   if not l_rs.eof then
	   l_vacnro = l_rs("vacnro")
   else
       l_vacnro = ""
   end if
end if

l_rs.close

'Busco los feriados y los almaceno en un arreglo

l_sql =         " SELECT * FROM feriado "
rsOpen l_rs, cn, l_sql, 0 

l_cantFeriados = 0
do until l_rs.eof 
	 if ((CInt(l_rs("tipferinro")) = 1) AND (CInt(l_rs("fericodext")) = 170 ) ) then
	    l_arrFeriados(l_cantFeriados) = l_rs("ferifecha")
		l_cantFeriados = l_cantFeriados + 1
	 end if
	 l_rs.movenext
loop
l_rs.Close


'Busco las siglas de los tipos de dias

l_sql =         " SELECT * FROM tipdia "
rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof 
     l_tipoDia(CInt(l_rs("tdnro"))) = mid(l_rs("tdsigla"),1,3)

	 l_rs.movenext
loop
l_rs.Close

Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

Set l_rs20 = Server.CreateObject("ADODB.RecordSet")

l_sqlTer =         " SELECT ternro FROM empleado WHERE empleg= " & leg
rsOpen l_rs20, cn, l_sqlTer, 0 

l_sqlemp = " SELECT ternro FROM empleado WHERE empreporta = " & l_rs20("ternro")

l_rs20.Close

l_sql = "SELECT emp_licnro,tdnro,elfechadesde,elfechahasta, elcantdias, terape, terape2, ternom, ternom2, empleg, empleado.ternro, licestnro "
l_sql = l_sql & " FROM emp_lic "
l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = emp_lic.empleado AND ((elfechadesde >=" & cambiafecha(l_desde,"YMD",true)
l_sql = l_sql & " and elfechahasta <= " & cambiafecha(l_hasta,"YMD",true) & ") "
l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(l_desde,"YMD",true)
l_sql = l_sql & " and elfechahasta <= " & cambiafecha(l_hasta,"YMD",true) 
l_sql = l_sql & " and elfechahasta >= " & cambiafecha(l_desde,"YMD",true) & ") "	
l_sql = l_sql & " or (elfechadesde >= " & cambiafecha(l_desde,"YMD",true)
l_sql = l_sql & " and elfechahasta >  " & cambiafecha(l_hasta,"YMD",true) 
l_sql = l_sql & " and elfechadesde <= " & cambiafecha(l_hasta,"YMD",true) & ") "	
l_sql = l_sql & " or (elfechadesde <  " & cambiafecha(l_desde,"YMD",true)
l_sql = l_sql & " and elfechahasta >  " & cambiafecha(l_hasta,"YMD",true) & ")) "
'l_sql = l_sql & " WHERE tdnro IN (2,3,4,5,7,18,19,23,25,31,34,35) " 'Licencias que se cargan por web
l_sql = l_sql & " WHERE tdnro IN (2,3,4,5,7,8,11,15,16,18,19,22,23,24,25,29,30,31,32,33,34,35) "
l_sql = l_sql & "   AND empleado.ternro IN (" & l_sqlemp & ")"
l_sql = l_sql & "   AND licestnro IN (1,2)"

if l_filtro <> "" then
   l_sql = l_sql & " AND " & l_filtro
end if

l_sql = l_sql & " ORDER BY " & l_orden

rsOpen l_rs, cn, l_sql, 0 

l_poner_titulo = true
l_pagina = 0

do until l_rs.eof

	'----------------------------------------------------------------------------------------
	if l_poner_titulo then
       l_poner_titulo = false
	   l_pagina = l_pagina + 1
	   l_nrolinea = 0
	   
	   l_descripcion = l_desde & " - " & l_hasta 
	   
	   response.write "<table border=0 cellpadding=0 cellspacing=0 style='background:#ffffff;' width='100%' >"
	   
       response.write "<tr><th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' colspan=" & (l_cant_dias + 7) & "><table class='TABLE2' border=0 cellpadding=0 cellspacing=0 style='background:#ffffff;' width='100%'>"	   
       response.write "<th colspan=" & (l_cant_dias + 7 ) & " align='center'><b style='font-size:11pt;'>Licencias Empleados</b>&nbsp;&nbsp;" & l_descripcion & "</th>"
       response.write "</table></td></tr>"
	   
	   'Imprimo los anio
       response.write "<tr>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' colspan='3'>Lic.&nbsp;Vac.</th>"	   
       'response.write "<th style='border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' colspan='3'>Lic.&nbsp;Tur.</th>"	   	   
	   
	   l_factual = CDate(l_desde)
	   l_i = 0
	   l_mesant = month(CDate(l_factual))
	   while DateDiff("d",l_factual,CDate(l_hasta)) >= 0
	       if month(CDate(l_factual)) <> l_mesant then
             response.write "<th style='border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' colspan=" & l_i & ">" & descMes(l_mesant) & "<br>"		   
		     l_i = 1
			 l_mesant = month(CDate(l_factual))
		   else
		     l_i = l_i + 1
		   end if

	       l_factual = DateAdd("d",1,l_factual)
	   wend

       response.write "<th style='border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' colspan=" & l_i & ">" & descMes(l_mesant) & "<br>"		   

       response.write "</tr>"
	   
	   'Imprimo los dias
       response.write "<tr>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Sal.</th>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Goz.</th>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Pen.</th>"
       'response.write "<th style='border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Sal.</th>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Goz.</th>"
       'response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Pen.</th>"
	   
	   l_factual = CDate(l_desde)
	   while DateDiff("d",l_factual,CDate(l_hasta)) >= 0
           response.write "<th style='background:" & colorDia(l_factual) & ";border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' align='center'>" & diaSemana(l_factual) & "</th>"

	       l_factual = DateAdd("d",1,l_factual)
	   wend
       response.write "</tr>"
	   
	   
	   l_nrolinea = l_nrolinea + 3
	   
	end if

	l_hay_datos = true
	
	l_emp_actual = l_rs("ternro")
	l_empleado   = l_rs("empleg") & "-" & l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
	l_salir = false
	
	inicializarDias
	
	do 

		if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(l_desde)) <= 0) and _
		   (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(l_hasta)) >= 0) then

		   cargarRango l_rs("elfechadesde"),l_rs("elfechahasta"),l_rs("tdnro"),l_rs("licestnro"),l_rs("emp_licnro")

		else
		   if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(l_desde)) > 0) and _
		      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(l_hasta)) >= 0) and _
		      (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(l_desde)) <= 0) then
			  
		      cargarRango l_desde,l_rs("elfechahasta"),l_rs("tdnro"),l_rs("licestnro"),l_rs("emp_licnro")
			  
		   else
		      if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(l_desde)) <= 0) and _
		         (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(l_hasta)) < 0)  and _
		         (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(l_hasta)) >= 0) then

		         cargarRango l_rs("elfechadesde"),l_hasta,l_rs("tdnro"),l_rs("licestnro"),l_rs("emp_licnro")

			  else
		         if (DateDiff("d",CDate(l_rs("elfechadesde")), CDate(l_desde)) > 0) and _
		            (DateDiff("d",CDate(l_rs("elfechahasta")), CDate(l_hasta)) < 0) then
					
		            cargarRango l_desde,l_hasta,l_rs("tdnro"),l_rs("licestnro"),l_rs("emp_licnro")
					
				 end if
			  end if
		   end if
		end if

	   l_rs.movenext
	   if l_rs.eof then
	      l_salir = true
	   else
	      l_salir = CStr(l_rs("ternro")) <> CStr(l_emp_actual) 
	   end if
	loop until l_salir

    'l_corresp1  = 0
    'l_cantidad1 = 0
    'l_total1    = 0
    'l_corresp2  = 0
    'l_cantidad2 = 0
    'l_total2    = 0
	
	'call buscarDatosVacaciones(l_vacnro,l_emp_actual,l_anio, l_corresp1, l_cantidad1)
	
	'l_total1 = l_corresp1 - l_cantidad1
	
	'call buscarDatosTurismo(l_vacnro,l_emp_actual,l_anio, l_corresp2, l_cantidad2)
	
	'l_total2 = l_corresp2 - l_cantidad2

    response.write "<tr>" 

	'response.write "<td nowrap style='font-size:7pt;border-bottom : thick solid 1;'>" & l_corresp1  & "</td>"
	'response.write "<td nowrap style='font-size:7pt;border-bottom : thick solid 1;'>" & l_cantidad1 & "</td>"
	'response.write "<td nowrap style='font-size:7pt;border-bottom : thick solid 1;'>" & l_total1    & "</td>"
	'response.write "<td nowrap style='border-left-style: solid;border-left-width: 1px;border-left-color: black;font-size:7pt;border-bottom : thick solid 1;'>" & l_corresp2  & "</td>"
	'response.write "<td nowrap style='font-size:7pt;border-bottom : thick solid 1;'>" & l_cantidad2 & "</td>"
	'response.write "<td nowrap style='border-right-style: solid;border-right-width: 1px;border-right-color: black;font-size:7pt;border-bottom : thick solid 1;'>" & l_total2    & "</td>"
	
	for l_i=1 to l_cant_dias
		   if l_arrDias(l_i) = 0 then
		      response.write "<td style='padding: 0 0 0 0;border-bottom : thick solid 1;'>&nbsp;</td>"
		   else
              if CInt(l_arrLicencias(l_i)) = 0 then
		         response.write "<td nowrap style='padding: 2 0 2 0;border-bottom : thick solid 1;'><div style='padding: 0 0 0 0;font-size:9px;background:" & obtColor(CLng(l_arrEstados(l_i))) & ";'>&nbsp;" & l_tipoDia(CInt(l_arrDias(l_i))) & "</div></td>"			  
			  else
		         response.write "<td nowrap style='cursor:hand; padding: 2 0 2 0;border-bottom : thick solid 1;' ondblclick='javascript:cambioEstado(" & l_arrLicencias(l_i) & "," & l_arrEstados(l_i) & ");'><div style='padding: 0 0 0 0;font-size:9px;background:" & obtColor(CLng(l_arrEstados(l_i))) & ";'>&nbsp;" & l_tipoDia(CInt(l_arrDias(l_i))) & "</div></td>"
			  end if
		   end if
	next
	
    response.write "</tr>"	
	
	l_nrolinea = l_nrolinea + 1
	
	if (l_nrolinea > l_Max_Lineas_X_Pag) AND (not l_rs.eof) then
	   l_poner_titulo = true
       l_nrolinea = 0		   	   
	   response.write "</table><p style='page-break-before:always'></p>"
	end if

loop

l_rs.Close

if l_hay_datos then

	response.write "</table>"
	
	response.write "<br><table cellpadding=3 cellspacing=0 style='visibility:hidden;'>"

	response.write "<tr>"
	response.write "<th align=center colspan=3 style='font-size:8pt;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Referencia</th>"
	response.write "</tr>"

	response.write "<tr>"
	response.write "<th style='font-size:7pt;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' align=center>Color</th>"
	response.write "<th style='font-size:7pt;border-left-style: solid;border-left-width: 1px;border-left-color: black;border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;' align=center>Descripci&oacute;n</th>"
	response.write "</tr>"
	
    response.write "<tr>"
	response.write "<td style='font-size:7pt;' align=center><div style='padding: 0 0 0 0;font-size:8px;background:" & Color_finSemana & ";'>&nbsp;&nbsp;</div></td>"
	response.write "<td style='font-size:7pt;' align=left>Sabados y Domingos</td>"
	response.write "</tr>" 	
	
    response.write "<tr>"
	response.write "<td style='font-size:7pt;' align=center><div style='padding: 0 0 0 0;font-size:8px;background:" & Color_feriado & ";'>&nbsp;&nbsp;</div></td>"
	response.write "<td style='font-size:7pt;' align=left>Feriados</td>"
	response.write "</tr>" 	
	
    response.write "<tr>"
	response.write "<td style='font-size:7pt;' align=center><div style='padding: 0 0 0 0;font-size:8px;background:" & obtColor(1) & ";'>&nbsp;&nbsp;</div></td>"
	response.write "<td style='font-size:7pt;' align=left>Pendiente</td>"
	response.write "</tr>" 	
	
    response.write "<tr>"
	response.write "<td style='font-size:7pt;' align=center><div style='padding: 0 0 0 0;font-size:8px;background:" & obtColor(2) & ";'>&nbsp;&nbsp;</div></td>"
	response.write "<td style='font-size:7pt;' align=left>Aprobada</td>"
	response.write "</tr>" 	
	
	response.write "</table><br>"
else
   response.write "No se encontraron datos."
end if
 
cn.Close

%>
</table>

<script>
  mover();
</script>

</body>
</html>

