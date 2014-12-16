<!--#include virtual="/turnos/shared/inc/antigfec.inc"-->
<% 
'-----------------------------------------------------------------------------------
'Archivo	: campcalc.inc
'Descripción: Include con funciones para campos calculados para Reportes 
'Autor		: Walter Otmen
'Fecha		: Septiembre 2003
'Modificado	: 
'	23-04-04 - Alvaro Bayon - Funciones adicionales
'------------------------------------------------------------------------------------

Dim l_rs_auxi
Dim l_sql_auxi
Dim l_valor_auxi
Set l_rs_auxi = Server.CreateObject("ADODB.RecordSet")

function activo_inactivo(valor)
	if CInt(valor) = -1 then
		activo_inactivo = "Activo"
	else
		activo_inactivo = "Inactivo"
	end if 
end function

function masc_fem(valor)
	if CInt(valor) = -1 then
		masc_fem = "Masculino"
	else
		masc_fem = "Femenino"
	end if
end function

function si_no(valor)
	if CInt(valor) = -1 then
		si_no = "Si"
	else
		si_no = "No"
	end if
end function

' Calcula la Edad de un Tercero
function Edad (ByVal tercero)
	l_sql_auxi = "SELECT terfecnac from tercero where ternro = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	if l_rs_auxi("terfecnac") <> "" then
		Edad = int((date - CDate(l_rs_auxi("terfecnac"))) / 365)
	end if
	l_rs_auxi.Close
end function

' Calcula la Antiguedad de un tercero
function Antiguedad (ByVal tercero)
		dim l_texto, l_dia, l_mes, l_anio, l_hab
	    call antigfec (tercero, date + 1, l_dia, l_mes, l_anio, l_hab)
		if l_anio = "" or l_anio = 0 then
			if l_mes = "" or l_mes = 0 then
				l_texto = l_dia & " día/s."
			else
				l_texto = l_mes & " mes/es " & l_dia & " día/s."
			end if	
		else
			l_texto = l_anio & " año/s " & l_mes & " mes/es " & l_dia & " día/s."	
		end if
		Antiguedad = l_texto
end function

' Ubica el Empleado al Cual Reporta
function ReportaA (ByVal tercero)
	dim l_texto, l_reportaa
	l_sql_auxi = "SELECT empreporta FROM v_empleado WHERE ternro = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	l_reportaa = l_rs_auxi(0)
	l_rs_auxi.Close
	if l_reportaa <> "" then
		l_sql_auxi = "SELECT empleg, v_empleado.terape, v_empleado.terape2, v_empleado.ternom, v_empleado.ternom2 from v_empleado where ternro = " & l_reportaa
		rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
		l_texto = l_rs_auxi(0) & " - " & l_rs_auxi(1) & " " & l_rs_auxi(2) & ", " & l_rs_auxi(3) & " " & l_rs_auxi(4)
		ReportaA = l_texto
		l_rs_auxi.Close
	else
		ReportaA = ""
	end if
end function

' Domicilio del Tercero
function Domicilio (ByVal tercero)
	dim l_texto
	l_sql_auxi = "SELECT calle, nro, piso, oficdepto FROM cabdom "
	l_sql_auxi = l_sql_auxi & "INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
	l_sql_auxi = l_sql_auxi & "WHERE domdefault = -1 AND ternro = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	if not l_rs_auxi.eof then
		l_texto = l_rs_auxi(0) & " " & l_rs_auxi(1) 
		if l_rs_auxi(2) <> "" then l_texto = l_texto & " Piso: " & l_rs_auxi(2)
		if l_rs_auxi(3) <> "" then l_texto = l_texto & " Dto: " & l_rs_auxi(3)
	else
		l_texto = ""
	end if
	l_rs_auxi.Close
	Domicilio = l_texto
end function

' Idiomas del Tercero
function Idiomalee (ByVal tercero)
	dim l_texto
	l_sql_auxi = "SELECT ididesc , idnivdesabr FROM emp_idi , idioma , idinivel "
	l_sql_auxi = l_sql_auxi & "WHERE emp_idi.idinro = idioma.idinro "
    l_sql_auxi = l_sql_auxi & "AND emp_idi.empidlee = idinivel.idnivnro "
	l_sql_auxi = l_sql_auxi & "AND empleado = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	l_texto = ""
	do until l_rs_auxi.eof
		l_texto = l_texto & l_rs_auxi(0) & " " & l_rs_auxi(1) & "<br>"
		l_rs_auxi.MoveNext
	loop
	l_rs_auxi.Close
	Idiomalee = l_texto
end function

function Idiomahabla (ByVal tercero)
	dim l_texto
	l_sql_auxi = "SELECT ididesc , idnivdesabr FROM emp_idi , idioma , idinivel "
	l_sql_auxi = l_sql_auxi & "WHERE emp_idi.idinro = idioma.idinro "
    l_sql_auxi = l_sql_auxi & "AND emp_idi.empidhabla = idinivel.idnivnro "
	l_sql_auxi = l_sql_auxi & "AND empleado = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	l_texto = ""
	do until l_rs_auxi.eof
		l_texto = l_texto & l_rs_auxi(0) & " " & l_rs_auxi(1) & "<br>"
		l_rs_auxi.MoveNext
	loop
	l_rs_auxi.Close
	Idiomahabla = l_texto
end function

function Idiomaescribe (ByVal tercero)
	dim l_texto
	l_sql_auxi = "SELECT ididesc , idnivdesabr FROM emp_idi , idioma , idinivel "
	l_sql_auxi = l_sql_auxi & "WHERE emp_idi.idinro = idioma.idinro "
    l_sql_auxi = l_sql_auxi & "AND emp_idi.empidescr = idinivel.idnivnro "
	l_sql_auxi = l_sql_auxi & "AND empleado = " & tercero
	rsOpen l_rs_auxi, cn, l_sql_auxi, 0 
	l_texto = ""
	do until l_rs_auxi.eof
		l_texto = l_texto & l_rs_auxi(0) & " " & l_rs_auxi(1) & "<br>"
		l_rs_auxi.MoveNext
	loop
	l_rs_auxi.Close
	Idiomaescribe = l_texto
end function


' Funcion concentradora de llamadas a funciones de campos calculados (Generar una entrada por cada funcion)
function CampCalc(ByVal TipoFuncion, ByVal campo, ByVal tercero)
Dim alfa
Dim parametros
if instr(TipoFuncion,"@") then
   	select Case campo
		Case "EDAD" CampCalc = Edad(tercero)
		Case "ANTIGUEDAD" CampCalc = Antiguedad(tercero)
           Case "REPORTAA" CampCalc = ReportaA(tercero)
           Case "DOMICILIO1" CampCalc = Domicilio(tercero)
		   Case "IDIOMALEE" CampCalc = Idiomalee(tercero)
		   Case "IDIOMAHABLA" CampCalc = Idiomahabla(tercero)
		   Case "IDIOMAESCRIBE" CampCalc = Idiomaescribe(tercero)
    end Select
else
	'En este caso "campo" es la función a aplicar sobre "tercero", que es un valor cualquiera.
	select case campo
		case "ACTIVO_INACTIVO"
			CampCalc = activo_inactivo(tercero)
		case "SI_NO"
			CampCalc = si_no(tercero)
		case "MASC_FEM"
			CampCalc = masc_fem(tercero)
		Case "FAMEDAD" 
			CampCalc = Edad(tercero) 
		case else
			CampCalc = campo
	end select
end if
end function

%>
