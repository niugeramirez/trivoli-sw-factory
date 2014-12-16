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
Modificado     : 07-10-2005 - Leticia A. - Adecuacion para que funcione desde Autogestion 
-->

<% 
on error goto 0

const Color_Feriado   = "#ccffff"
const Color_FinSemana = "#e0e0e0"
const Color_Aprobado  = "#ccffcc"
const Color_Pendiente = "#edbf6b"

Const l_Max_Lineas_X_Pag = 10000
const l_nro_col = 4

' Variables
Dim l_ternro
Dim leg
dim l_rs
dim l_rs20
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

Dim l_rs2

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
 
 l_factual = CDate(l_desde)
 l_i = 1
 while DateDiff("d",l_factual,CDate(l_hasta)) >= 0
 
     if (DateDiff("d",CDate(desde),l_factual) >= 0) AND (DateDiff("d",l_factual,CDate(hasta)) >= 0) then
	    l_arrDias(l_i) = tdnro
		l_arrEstados(l_i) = estado 
	    l_arrLicencias(l_i) = licencia
	 end if

	 l_i = l_i + 1
     l_factual = DateAdd("d",1,l_factual)
 wend

end sub 'cargarRango(desde,hasta,tdnro)

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
  parent.ifrm2.document.body.scrollTop = document.body.scrollTop;
}

</script>

</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onScroll="mover();">
<%

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

l_sqlTer = " SELECT ternro FROM empleado WHERE empleg= " & leg
rsOpen l_rs20, cn, l_sqlTer, 0 

l_sqlemp = " SELECT ternro FROM empleado WHERE empreporta = " & l_rs20("ternro")

l_rs20.Close


l_sql = "SELECT emp_licnro,tdnro,elfechadesde,elfechahasta, elcantdias, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2, tercero.ternomhab, empleg, empleado.ternro, licestnro "
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
l_sql = l_sql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro"
'l_sql = l_sql & " WHERE tdnro IN (3,4,7,19,2,18,35,25,23,31,5,34) "  ' Licencias que se cargan por web
l_sql = l_sql & " WHERE empleado.empest = -1 and empleado.ternro IN (" & l_sqlemp & ")"
l_sql = l_sql & "   AND licestnro IN (1,2)"

if l_filtro <> "" then
   l_sql = l_sql & " AND " & l_filtro
end if

l_sql = l_sql & " ORDER BY empleado." & l_orden

rsOpen l_rs, cn, l_sql, 0 

response.write l_sql
 reponse.end

l_poner_titulo = true
l_pagina = 0

do until l_rs.eof

	'----------------------------------------------------------------------------------------
	if l_poner_titulo then
       l_poner_titulo = false
	   l_pagina = l_pagina + 1
	   l_nrolinea = 0
	   
	   l_descripcion = l_desde & " - " & l_hasta 
	   
	   response.write "<table o border=0 cellpadding=0 cellspacing=0 style='background:#ffffff;' width='100%'>"
	   
       response.write "<tr><th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'><table class='TABLE2' border=0 cellpadding=0 cellspacing=0 style='background:#ffffff;' width='100%'>"	   
       response.write "<th colspan=1 align='center'><b style='font-size:11pt;'>&nbsp;</th>"
       response.write "</table></td></tr>"
	   
	   'Imprimo los anio
       response.write "<tr>"
       response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>Meses</th>"
       response.write "</tr>"
	   
	   'Imprimo los dias
       response.write "<tr>"
       response.write "<th style='border-bottom-style: solid;border-bottom-width: 1px;border-bottom-color: black;'>D&iacute;as</th>"
       response.write "</tr>"
	   
	   l_nrolinea = l_nrolinea + 3
	   
	end if

	l_hay_datos = true
	
	l_emp_actual = l_rs("ternro")
	l_empleado   = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
'	l_empleado   = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternomhab")
	l_salir = false

	do 

	   l_rs.movenext
	   if l_rs.eof then
	      l_salir = true
	   else
	      l_salir = CStr(l_rs("ternro")) <> CStr(l_emp_actual) 
	   end if

	loop until l_salir

    response.write "<tr>" 

	response.write "<td nowrap style='font-size:7pt;border-bottom : thick solid 1;'>" & l_empleado  & "</td>"

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
	
	response.write "<br><table cellpadding=3 cellspacing=0>"

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
</body>
</html>

