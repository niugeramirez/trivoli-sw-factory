<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : ag_evaluar_eventos_cap_07.asp
Descripcion    : Muestra los Resultados de las Evaluaciones
Autor		   : Raul Chinestra
Fecha		   : 12/07/2007
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim Graf

Dim l_tesnro
Dim l_fornro
Dim l_tesfin
Dim l_ternro
Dim l_tesfec
Dim l_teshor
Dim l_testie
Dim l_tercero
Dim l_RespondidasOk
Dim l_RespondidasMal
Dim l_totalpreguntas
Dim l_Ponderacion
Dim l_TotalPonderacion
Dim l_ttespond
Dim l_ttesmin
Dim l_ttestie

l_tesnro	= request.QueryString("tesnro")
l_fornro	= request.QueryString("fornro")
l_tesfin	= request.QueryString("tesfin")

l_ternro = l_ess_ternro

Set l_rs = Server.CreateObject("ADODB.RecordSet")

function GrafDial(l_minimo,l_requerido,l_maximo,l_valor)
'Genera el grafico
'l_minimo = valor inicial del graf
'l_requerido = valor hasta el cual va la zona roja y a partir de este, hasta maximo la zona es verde
'l_maximo = valor final
'l_valor = valor que va a mostrar la aguja
	dim strChartXML
	
	strChartXML = "<Chart upperLimit='" & l_maximo & "' lowerLimit='" & l_minimo & "' majorTMNumber='11' majorTMHeight='10' minorTMNumber='9' minorTMHeight='3' pivotRadius='0' majorTMThickness='2' showGaugeBorder='0' gaugeOuterRadius='110' gaugeOriginX='175' gaugeOriginY='170' gaugeScaleAngle='180' gaugeInnerRadius='2' numberScaleValue='1000,1000' numberScaleUnit='K,M' formatNumberScale='1' displayValueDistance='30' decimalPrecision='0' tickMarkDecimalPrecision='2'>"
	strChartXML = strChartXML & "<colorRange>"
	strChartXML = strChartXML & "<color minValue='" & l_minimo & "' maxValue='" & l_requerido & "' code='B41527'/>"
	strChartXML = strChartXML & "<color minValue='" & l_requerido & "' maxValue='" & l_maximo & "' code='399E38'/>"
	strChartXML = strChartXML & "</colorRange>"
	strChartXML = strChartXML & "<dials>"
	strChartXML = strChartXML & "<dial value='" & l_valor & "' borderAlpha='0' bgColor='000000' baseWidth='10' topWidth='1' radius='105' />"
	strChartXML = strChartXML & "</dials>"
	strChartXML = strChartXML & "<customObjects>"
	strChartXML = strChartXML & "<objectGroup xPos='175' yPos='172'>"
	strChartXML = strChartXML & "<object type='circle' xPos='0' yPos='2.5' radius='118' startAngle='0' endAngle='180' fillPattern='linear' fillAsGradient='1' fillColor='dddddd,666666' fillAlpha='100,100' fillRatio='50,50' fillDegree='0' showBorder='1' borderColor='444444' borderThickness='2'/>"
	strChartXML = strChartXML & "<object type='circle' xPos='0' yPos='0' radius='115' startAngle='0' endAngle='180' fillPattern='linear' fillAsGradient='1' fillColor='666666,ffffff' fillAlpha='100,100' fillRatio='50,50' fillDegree='0'/>"
	strChartXML = strChartXML & "</objectGroup>"
	strChartXML = strChartXML & "<objectGroup xPos='175' yPos='172' showBelowChart='0'>"
	strChartXML = strChartXML & "<object type='circle' xPos='0' yPos='4' radius='15' startAngle='0' endAngle='180' color='000000'/>"
	strChartXML = strChartXML & "<object type='circle' xPos='0' yPos='4' radius='11' startAngle='0' endAngle='180' color='7F7F7F'/>"
	strChartXML = strChartXML & "<object type='circle' xPos='0' yPos='4' radius='3' startAngle='0' endAngle='180' color='ffffff'/>"
	strChartXML = strChartXML & "</objectGroup>"
	strChartXML = strChartXML & "</customObjects>"
	strChartXML = strChartXML & "</Chart>"
	
	GrafDial=strChartXML
end function


'-------------------------------------------------------------------------------------------------------------------------
' Busco el Total de las Preguntas del Formulario
'-------------------------------------------------------------------------------------------------------------------------
function TotalPreguntas()
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT count(*)  " 
	l_sql = l_sql & " FROM pos_pregunta "
	l_sql = l_sql & " WHERE fornro = " & l_fornro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		TotalPreguntas = l_rs(0)
	else
		TotalPreguntas = 0
	end if
	l_rs.Close
	Set l_rs = Nothing
end function 

function ControlarEvaluacion()

	Dim l_rs
	Dim l_rs2
	Dim l_sql
	Dim l_opcOK
	Dim l_opcRes

	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "SELECT * " 
	l_sql = l_sql & " FROM pos_pregunta "
	l_sql = l_sql & " WHERE fornro = " & l_fornro
	rsOpen l_rs, cn, l_sql, 0 
	
	l_RespondidasOk = 0
	l_RespondidasMal = 0
	l_Ponderacion = 0
	l_TotalPonderacion = 0
	
	do while not l_rs.eof
	
		l_TotalPonderacion = l_TotalPonderacion + clng(l_rs("prepond"))
		'--------------------------------------------------------------
		' Busco la Opción Correcta
		'--------------------------------------------------------------
		l_sql = "SELECT * " 
		l_sql = l_sql & " FROM pos_opcion "
		l_sql = l_sql & " WHERE prenro = " & l_rs("prenro")
		l_sql = l_sql & " AND opcOK = -1 "
		rsOpen l_rs2, cn, l_sql, 0 
		if not l_rs2.eof then 
			l_opcOK = l_rs2("opcnro")
		else 
			l_opcOK = -1
		end if 
		l_rs2.close

		'--------------------------------------------------------------
		' Busco la Opción Contestada en el Test por el Empleado
		'--------------------------------------------------------------
		l_sql = "SELECT * " 
		l_sql = l_sql & " FROM pos_respuesta "
		l_sql = l_sql & " WHERE tesnro = " & l_tesnro
		l_sql = l_sql & " AND prenro = " & l_rs("prenro")
		rsOpen l_rs2, cn, l_sql, 0 
		if not l_rs2.eof then 
			l_opcRes = l_rs2("resval")
		else 
			l_opcRes = -1
		end if 
		l_rs2.close
		
		if l_opcOK = l_opcRes then 
			l_RespondidasOk = l_RespondidasOk + 1
			l_Ponderacion = l_Ponderacion + clng(l_rs("prepond"))
		else
			l_RespondidasMal = l_RespondidasMal + 1
		end if 	
		
		l_rs.movenext
	loop
	
	l_rs.Close
	Set l_rs = Nothing
end function

'------------------------------------------------------------------------------------------------------------------------
' Busco los datos del Test
'------------------------------------------------------------------------------------------------------------------------
Sub DatosTest()

	Dim l_rs
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "SELECT * " 
	l_sql = l_sql & " FROM test "
	l_sql = l_sql & " INNER JOIN pos_tipotest ON pos_tipotest.ttesnro = test.ttesnro "
	l_sql = l_sql & " WHERE tesnro = " & l_tesnro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then 
		l_tesfec   = l_rs("tesfec")
		l_teshor   = l_rs("teshor")
		l_testie   = l_rs("testie")
		l_ttespond = l_rs("ttespond")
		l_ttesmin  = l_rs("ttesmin")
		l_ttestie  = l_rs("ttestie")
	end if
end sub

'------------------------------------------------------------------------------------------------------------------------
' Busco los datos del Empleados
'------------------------------------------------------------------------------------------------------------------------

Sub DatosEmpleado

	Dim l_rs
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "SELECT terape, terape2, ternom, ternom2  " 
	l_sql = l_sql & " FROM tercero "
	l_sql = l_sql & " WHERE ternro = " & l_ternro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_tercero	= l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
	end if
	l_rs.close
	set l_rs = nothing
end sub


ControlarEvaluacion
DatosTest
DatosEmpleado
l_totalpreguntas = TotalPreguntas()

'-------------------------------------------------------------------------------------------------------------------
'Funcion que genera los graficos
'Buscar los datos de los parametros de acuerda a las correspondientes sql
'-------------------------------------------------------------------------------------------------------------------
Graf = GrafDial(0,l_ttespond,l_TotalPonderacion,l_Ponderacion)

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title></title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table cellpadding="0" cellspacing="0" border="0">

<tr> 
	<td colspan="4" align="center"><b>Resultados de la Evaluación</b></td>	
</tr>
<tr> 
	<td colspan="4">&nbsp;</td>	
</tr>

<tr> 
	<td width="25%">Empleado: <b><%= l_tercero %></b></td>
	<td width="25%">&nbsp;</td>
	<td width="25%" align="center">Fecha: <%= l_tesfec %></td>
	<td width="25%" align="center">Hora: <%= l_teshor %></td>	
</tr>

<% if l_ttestie = -1 then %>
<tr> 
	<td width="25%">Tiempo Límite del Test: <%= l_ttesmin %> Min.</td>
	<td width="25%">&nbsp;</td>
	<td>Tiempo Evaluación:&nbsp;<%= l_testie%></td>	
	<td width="25%">&nbsp;</td>		
</tr>
<% end if %>
<tr> 
	<td align="right">Total de Preguntas:&nbsp;<br>
    	Correctas:&nbsp;<br>
		Incorrectas:&nbsp;<br><br><br>
		Total Ponderacion:&nbsp;<br>
		Ponderación Obtenida:&nbsp;
	</td>	
	<td align="left">
	<%= l_totalpreguntas%><br>
   	<%= l_RespondidasOK%><br>
	<%= l_RespondidasMal%><br><br><br>
	<%= l_TotalPonderacion%><br>
	<%= l_Ponderacion%>
	</td>
	<td colspan="2" rowspan="3">
	<div align="center" class="text"> 
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="350" height="200">
	<PARAM NAME="FlashVars" value="&dataXML=<%=Graf%>">
	<param name="movie" value="../../shared/Charts/FI2_Angular.swf?chartWidth=350&chartHeight=200">
	<param name="quality" value="high">
   	</object>
	</div>
	</td>
</tr>
</table>
</body>
</html>
<% Set l_rs = Nothing %>