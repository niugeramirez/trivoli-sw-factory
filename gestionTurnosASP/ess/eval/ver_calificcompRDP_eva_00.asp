<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: ver_gralobj_eva_00.asp
'Descripción	: Cargar resultado gral de objetivos
'Autor			: 04-01-2005
'Fecha			: CCRossi
'Modificado		: 04-02-2005 L Amadio
' 				: 19-05-2005 - LA - cambio de disposicion de la info.
'				: 21-07-2005 - L.A. - permitir evaluar RDP si existe al menos una RDE cerrada
'				: 03-08-2005 - L.A. - Cambiar cod de proyecto por cod de evento.
'            	  13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				  24/05/2007 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================

on error goto 0

' Variables
 
' de uso local  
  dim l_evatrnro 
  dim l_evacabnro 
  dim l_ternro  
  dim l_evaluador 
  dim l_evaevenro
  dim l_proyectos
  dim l_cantproys
  dim l_cantproysaprob
  dim l_proys
  dim l_datos 
  dim l_sincerrarRDE 
  
  dim l_gerente
  dim l_socio
  dim l_horas
  dim l_objResu
  dim l_objGrabar
  dim l_objGrabar2
  
  dim l_areaResu
  dim l_areaDescrip
  dim l_areaGrabar
  dim l_areaGrabar2
  dim l_compResu
  dim l_compDescrip
  dim l_compGrabar
  
  Dim l_evafacnro
  Dim l_evatitdesabr
  Dim l_observables 
  Dim l_interpretaciones

  dim l_areatrnro
  dim l_evaareadesc

 dim i
 Dim l_terminarsecc
 dim l_estrnro 
 dim l_evatevnro
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs2
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  Dim l_evapernro  ' VERRR si se pasa siempre este parametro

  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")
  l_evapernro  = Request.QueryString("evapernro")

   
'--------------------------------------------------------------------------------
Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT evacab.evacabnro, evatevnro, evaevento.evaevenro  " 
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
l_sql = l_sql & " WHERE evldrnro= " & l_evldrnro 
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then
	l_evacabnro = l_rs1("evacabnro")
	l_evatevnro = l_rs1("evatevnro")
	l_evaevenro = l_rs1("evaevenro")
	l_evatevnro = l_rs1("evatevnro")
end if  
l_rs1.Close
set l_rs1 = nothing
  
  
' buscar el ternro del EVALUADO (Aconsejado----------------------------------------------------
l_ternro=""
l_estrnro="0"
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado  "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " WHERE evadetevldor.evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_ternro = l_rs("empleado")
end if	
l_rs.Close
set l_rs=nothing

 'buscar las estructuras ACTIVAS del empleado -----------------------------------------------------------------
if trim(l_ternro) <> "" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT his_estructura.estrnro, teorden "
	l_sql = l_sql & " FROM his_estructura "
	l_sql = l_sql & " INNER JOIN estructura		ON estructura.estrnro=his_estructura.estrnro "
	l_sql = l_sql & " INNER JOIN tipoestructura ON tipoestructura.tenro=estructura.tenro "
	l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
	l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
	l_sql = l_sql & " ORDER BY teorden " 
	rsOpen l_rs, cn, l_sql, 0
	do while not l_rs.eof 
		l_estrnro = l_estrnro & "," & l_rs("estrnro")
		l_rs.MoveNext
	loop	
	l_rs.Close
	set l_rs=nothing
end if


' _________________________________________________________________________________
' buscar todos los proyectos en que participo el empleado (igual periodo y estrnro que evento RDP - RDE cerrada)
l_proyectos = "0"
l_cantproys = 0
l_proys=""
l_cantproysaprob = 0
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaproyecto.evaproynro, cabaprobada, evento.evaevenro "
l_sql = l_sql & " FROM evaevento "
l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evapernro = evaevento.evaperact "
l_sql = l_sql & " INNER JOIN evaevento evento ON evento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro  AND evatip_estr.tenro ="& cdepartamento 
l_sql = l_sql & " WHERE  evaevento.evaevenro =" &l_evaevenro &" AND evacab.empleado="&l_ternro
		' evacab.cabaprobada= -1 AND
rsOpen l_rs1, cn, l_sql, 0
do while not l_rs1.eof
	if l_rs1("cabaprobada") = -1 then
		l_proyectos = l_proyectos & "," & l_rs1("evaproynro") 
		l_cantproysaprob =l_cantproysaprob +1 
	else
		l_proys =l_proys &  " - " & l_rs1("evaevenro") ' l_rs1("evaproynro")
	end if
	l_cantproys= l_cantproys +1 
l_rs1.MoveNext
loop
l_rs1.close

l_proyectos = Split(l_proyectos,",")

l_sincerrarRDE="NO"
if l_cantproys <> l_cantproysaprob then
	l_sincerrarRDE="SI"
end if


'_______________________________________________________________________________
' fijarse que halla posibles resultados configurado, para la seccion
'  pero los posibles resultados se definen para cada competencia
' Crear registros de evaresultado para los facnro y el evldrnro (para evaluado y evaluador)
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT Distinct evaseccfactor.evafacnro, evatitulo.evatitnro, evaseccfactor.evaseccnro,  " ' evaresu.evatrnro, 
l_sql = l_sql & "  evatitulo.evatitdesabr "
l_sql = l_sql & " FROM evaseccfactor "
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
l_sql = l_sql & " ORDER BY evatitulo.evatitnro "
rsOpen l_rs, cn, l_sql, 0

set l_cm = Server.CreateObject("ADODB.Command")
if not l_rs.eof then
  	l_evatitdesabr = ""
	do while not l_rs.eof
		if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then
			l_evatrnro  = "null"
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT *  FROM  evaarea "
			l_sql = l_sql & " WHERE evaarea.evldrnro   = " & l_evldrnro
			l_sql = l_sql & " AND   evaarea.evatitnro  = " & l_rs("evatitnro")
			rsOpen l_rs1, cn, l_sql, 0
			if l_rs1.EOF and l_evatevnro=cint(cconsejero) then
				l_sql = "INSERT INTO evaarea "
				l_sql = l_sql & " (evldrnro, evatitnro, evatrnro, evaareadesc) "
				l_sql = l_sql & " VALUES ("& l_evldrnro & "," & l_rs("evatitnro")& ","& l_evatrnro & ",'')"
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
			end if
			l_rs1.Close
			set l_rs1=nothing
			l_evatitdesabr = l_rs("evatitdesabr")
		end if
		
		l_evafacnro = l_rs("evafacnro")
		l_evatrnro  = "null"
	  	
  		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evadetevldor.evldrnro, evaresultado.evldrnro as evldrnroresu  "
		l_sql = l_sql & " FROM evacab "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & " LEFT JOIN evaresultado ON evadetevldor.evldrnro = evaresultado.evldrnro AND evaresultado.evafacnro = "& l_rs("evafacnro")
		l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro & " AND evadetevldor.evaseccnro= "& l_evaseccnro
		rsOpen l_rs1, cn, l_sql, 0
		do while not l_rs1.EOF 
			if isNull(l_rs1("evldrnroresu")) then 
				l_sql = "INSERT INTO evaresultado "
				l_sql = l_sql & " (evldrnro, evafacnro, evatrnro, evaresudesc) "
				l_sql = l_sql & " VALUES (" & l_rs1("evldrnro") & "," & l_rs("evafacnro") & ","& l_evatrnro & ",'')"
				l_cm.activeconnection = Cn 
				l_cm.CommandText = l_sql   
				cmExecute l_cm, l_sql, 0   
			end if
		l_rs1.MoveNext
		loop
		l_rs1.Close
		set l_rs1=nothing
	l_rs.MoveNext
	loop
end if
l_rs.Close
set l_rs=nothing


'_______________________________________________________________________
sub datosProyecto (proynro, datos)
dim evento

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 l_sql = " SELECT evaevenro FROM evaevento WHERE evaproynro="& proynro
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
 	evento = l_rs("evaevenro")
 end if
 l_rs.Close
 
 l_sql = "SELECT evaengage.evaengnro, evaengdesabr, evaclinom, proygerente, proysocio, evaproyfdd, evaproyecto.evaproynro"
 l_sql = l_sql & " FROM evaengage "
 l_sql = l_sql & " INNER JOIN evacliente  ON evacliente.evaclinro  = evaengage.evaclinro "
 l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evaengnro = evaengage.evaengnro "
 l_sql = l_sql & " INNER JOIN evaproyemp  ON evaproyemp.evaproynro = evaproyecto.evaproynro "
 l_sql = l_sql & " WHERE evaproyemp.ternro = " & l_ternro & " AND evaproyecto.evaproynro=" & proynro
 		' XXXXXXXXXXX y por estructura tendria que preguntar??????????'
 rsOpen l_rs, cn, l_sql, 0 

 if not l_rs.eof then
 		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		' buscar nom de gerente y socio
		l_sql= " SELECT terape, terape2,ternom, ternom2 "
		l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs("proygerente")
  		rsOpen l_rs1, cn, l_sql, 0 
		l_gerente = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
		l_rs1.Close 
		l_sql= " SELECT terape, terape2,ternom, ternom2 "
		l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs("proysocio")
  		rsOpen l_rs1, cn, l_sql, 0
		l_socio = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
		l_rs1.Close
		
		l_sql = "SELECT evadetevldor.evldrnro, horas  " 
		l_sql = l_sql & " FROM evaproyecto " 
		l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & " INNER JOIN evadatosadm ON evadatosadm.evldrnro = evadetevldor.evldrnro "
		l_sql = l_sql & " WHERE evacab.empleado =" & l_ternro & " AND evaproyecto.evaproynro=" & l_rs("evaproynro") & " AND evadetevldor.evatevnro=" & cevaluador
		rsOpen l_rs1, cn, l_sql, 0 
		if l_rs1.eof then 
			l_horas = "--"
		else
			l_horas = l_rs1("horas")
		end if
		l_rs1.Close
		
		' Datos grales del engagement
		' datos = "Proyecto: " & proynro & "<br>"
		datos = "Evento: " & evento & "<br>"
		datos = datos & "Engagement:" & l_rs("evaengdesabr")& "<br>"
		datos = datos & "Cliente:" & l_rs("evaclinom")& "<br>"
		datos = datos & "Gerente:" & l_gerente & "<br>"
		datos = datos & "Socio:" & l_socio & "<br>"
		datos = datos & "Horas Imputadas: " & l_horas & "<br>"
		datos = datos & "Fecha inicio: " & l_rs("evaproyfdd")& "<br>"
 end if
 l_rs.close
 
end sub
  
 
  '-------------------------------------------------------------------
' _____________________________________________________________________
sub datosArea (areaResu, areaDescrip, areaGrabar, areaGrabar2)
areaResu=""
areaDescrip=""

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evatrnro, evaareadesc  "
l_sql = l_sql & " FROM evaarea "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaarea.evldrnro"
l_sql = l_sql & " WHERE evaarea.evatitnro  = " & l_rs("evatitnro")
l_sql = l_sql & " AND   evadetevldor.evacabnro = " & l_evacabnro & " AND evadetevldor.evatevnro="& cconsejero
'l_sql = l_sql & " AND   evaarea.evldrnro  = " & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0
if not l_rs1.eof then 
	l_areatrnro=l_rs1("evatrnro")
	l_evaareadesc=l_rs1("evaareadesc")
end if
l_rs1.close
set l_rs1=nothing

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaresu.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr  "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaresu.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
l_sql = l_sql & " AND   evatitulo.evatitnro = " & l_rs("evatitnro")
l_sql = l_sql & " order by evatrvalor "
rsOpen l_rs1, cn, l_sql, 0

areaResu= areaResu & "<select name=areatrnro"& l_rs("evatitnro")&" disabled style="" width:120 "">"
areaResu= areaResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
do while not l_rs1.eof
	areaResu= areaResu & "<option value="&l_rs1("evatrnro")&">&nbsp;&nbsp;"& l_rs1("evatrdesabr")&"</option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
areaResu= areaResu & "	</select> "
areaResu= areaResu & "  <script>document.datos.areatrnro"&l_rs("evatitnro")&".value='"&l_areatrnro&"';</script>"

areaGrabar="<a href=# onclick=""return false;"">Grabar</a>"	
areaGrabar2= "<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabado"&l_rs("evatitnro")&"""  size=""1"">"
end sub   

' ____________________________________________________________________________
' ____________________________________________________________________________
sub datosCompetenc (compResu,interpr, compDescrip, compGrabar, evafacnro, evatevnro, evatrnro) 'buscar evaresu
compResu = ""
compDescrip = ""
interpr=""
	
		' ARREGLARRRRRRRRRRRRRRRRRRRRRR -- > ver si se puede cambiar segun el rol? y la competencia???
Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT distinct evaresu.evatrnro,evaresu.evaresudes, evatipresu.evatrvalor, evatipresu.evatrdesabr " '
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro 
l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rs("evafacnro") & " AND NOT (evaresudes Is Null) "
l_sql = l_sql & " order by evaresu.evatrnro "
rsOpen l_rs1, cn, l_sql, 0 
do while not l_rs1.eof
	if trim(l_rs1("evaresudes")) <>"" and not isnull(l_rs1("evaresudes")) then
		interpr = interpr & l_rs1("evatrdesabr") &": "&unescape(l_rs1("evaresudes")) & "\n"
	end if
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing


Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT  evaresu.evatrnro, evaresu.evaresudes, "
l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro  & " AND   evaresu.evafacnro  = " & evafacnro
l_sql = l_sql & " order by evaresu.evatrnro "
rsOpen l_rs1, cn, l_sql, 0 

compResu = compResu & "<select name=evatrnro"&evafacnro&evatevnro&" style="" width:100 "" disabled>"
compResu = compResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
do while not l_rs1.eof
	compResu = compResu & "<option value="& l_rs1("evatrnro")&">"& "&nbsp;"&l_rs1("evatrdesabr")&" </option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
compResu = compResu & "</select> "
compResu = compResu & " <script>document.datos.evatrnro"&evafacnro&evatevnro&".value='"& evatrnro&"'</script>"
compDescrip = "	<input type=hidden name=evaresudesc"&evafacnro&evatevnro&">"

compGrabar= "<a href=# >Grabar</a>"  ' CAMPO=GRABAR ????
compGrabar = compGrabar & "<br>" &"<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabado"&evafacnro&evatevnro&""" size=""1"">"		

if trim(interpr)= "" then
	interpr = "<a href=# onclick=""alert('No hay Interpretaciones cargadas para estos resultados.'); "">?</a>"
else
	interpr= "<a href=# onclick=""alert('"& interpr & "'); "">?</a>"
end if
 end sub

 '____________________________________________________
 '¡ ___________________________________________________
sub	datosObservables(obs)
obs = ""
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT evadescomp.evadcdes, estructura.estrdabr "
l_sql = l_sql & " FROM evadescomp "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evadescomp.estrnro "
l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs("evafacnro")
'l_sql = l_sql & " AND   evadescomp.tenro     =  " & ctenro
l_sql = l_sql & " AND   evadescomp.estrnro IN (" & l_estrnro & ")"
rsOpen l_rs1, cn, l_sql, 0

l_observables=""
do while not l_rs1.eof
	obs = obs & l_rs1("estrdabr") & " - "& l_rs1("evadcdes")& "\n"
	l_rs1.MoveNext
loop
l_rs1.Close
set l_rs1 = nothing

if trim(l_observables)="" then
	l_observables="<a href=# onclick=""alert('No hay definidas Conductas Observables \n para las Estructuras del Empleado \n y la Competencia.')"">?</a>"
else	
	l_observables="<a href=# onclick=""alert('"& l_observables&"')"">?</a>"
end if
end sub
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<style>
.rev {
	font-size: 10;
	border-style: none;
}
</style>
</head>

<script>

/* function Controlar(texto,valor){
	if (texto.value==""){
		alert('Ingrese un Objetivo.');
		texto.focus();		return false;
	}else
		if (valor.value==""){
			alert('Seleccione un resultado.');
			valor.focus();			return false;
		}	else	return true;
}	*/

function Controlar(resu){
	if ((resu.value=="")||(resu.value=="0")){
		alert('Seleccione un resultado.');
		resu.focus();
		return false;
	}
	else
		return true;
}	

function Blanquear(texto){
 var  aux;
 aux = replaceSubstring(texto,"'","")
 aux = replaceSubstring(aux,"´","")
 return aux;
}

function replaceSubstring(inputString, fromString, toString) {
   // Goes through the inputString and replaces every occurrence of fromString with toString
   var temp = inputString;
   if (fromString == "") {
      return inputString;
   }
   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
      var midStrings = new Array("~", "`", "_", "^", "#");
      var midStringLen = 1;
      var midString = "";
      // Find a string that doesn't exist in the inputString to be used
      // as an "inbetween" string
      while (midString == "") {
         for (var i=0; i < midStrings.length; i++) {
            var tempMidString = "";
            for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
            if (fromString.indexOf(tempMidString) == -1) {
               midString = tempMidString;
               i = midStrings.length + 1;
            }
         }
      } // Keep on going until we build an "inbetween" string that doesn't exist
      // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + midString + toTheRight;
      }
      // Next, replace the "inbetween" string with the "toString"
      while (temp.indexOf(midString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(midString));
         var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } // Ends the check to see if the string being replaced is part of the replacement string or not
   return temp; // Send the updated string back to the user
} // Ends the "replaceSubstring" function


</script>
<body leftmargin="0" topmargin="0" rightmargin="0">
<form name="datos">
<input type="Hidden" name="terminarsecc" value="--">
<input type="Hidden" name="terminarsecc2" value="">

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<% if l_sincerrarRDE="SI" then %>
<tr height="20">
	<td colspan="<%= 6 + Ubound(l_proyectos)*2%>" align="left" width="25%">
		<% if l_cantproysaprob > 0 then  %>
		<b> AVISO:</b> El empleado no tiene todas sus RDE's cerradas. <br>
		<% else %>
		<b> AVISO:</b> No se permite calificar la secci&oacute;n, dado que el empleado no tiene ninguna RDE cerrada. <br>		  
		<% end if%>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; Eventos sin RDE's cerradas: <%=l_proys%>
	</td>
</tr>
<tr><td colspan="<%= 6 + Ubound(l_proyectos)*2%>" align="center">&nbsp;</td> </tr>
<% end if %>
<tr height="20">
	<td colspan="<%= 6 + Ubound(l_proyectos)*2%>" align="center" width="25%">
		<b>CALIFICACI&Oacute;N POR COMPETENCIAS RDP&rsquo;s</b>
	</td>
</tr>
<tr><td colspan="<%= 6 + Ubound(l_proyectos)*2%>" align="center">&nbsp;</td> </tr>
<tr>
	<th rowspan="2" nowrap class="th2">AREAS/COMPETENCIAS </th>
	<th rowspan="2" colspan="2" class="th2">EVALUADO</th>
	<th rowspan="2" colspan="2" class="th2">CONSEJERO</th>
	<th rowspan="2" class="th2">OBS. </th>
	<%for i= 1 to Ubound(l_proyectos) 
		datosProyecto l_proyectos(i), l_datos %>
		<th colspan="2" align="left" nowrap class="th2"> <%=l_datos%></th>
	<% next %>
</tr>
<tr>
	<% for i= 1 to Ubound(l_proyectos) %>
		<th class="th2"> Evaluado</th>
		<th class="th2"> Revisor</th>
	<% next %>
</tr>
<tr><td colspan=<%=6 + Ubound(l_proyectos)*2%>>&nbsp;</td></tr>
<% 'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT distinct evaresultado.evldrnro, evaresultado.evafacnro,evaresultado.evatrnro " 'evaresultado.evaresudesc
   l_sql = l_sql & " ,evafactor.evafacdesabr, evafactor.evafacdesext, evaseccfactor.orden "
   l_sql = l_sql & " ,evatitulo.evatitnro, evatitulo.evatitdesabr, evadetevldor.evatevnro, evaoblieva.evaobliorden "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " INNER JOIN evaoblieva    ON evaoblieva.evaseccnro=evaseccfactor.evaseccnro "
   'l_sql = l_sql & " INNER JOIN evaresultado  ON evadetevldor.evldrnro = evaresultado.evldrnro "
   l_sql = l_sql & " LEFT JOIN evadetevldor  ON evadetevldor.evldrnro = evaresultado.evldrnro AND evadetevldor.evatevnro=evaoblieva.evatevnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =" & l_evaseccnro
   l_sql = l_sql & "   AND evadetevldor.evacabnro="& l_evacabnro
   l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evaseccfactor.orden, evaoblieva.evaobliorden"
   
   l_evatitdesabr="" 
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof 
		if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then
			datosArea l_areaResu, l_areaDescrip, l_areaGrabar, l_areaGrabar2  'buscar evaarea 
%>
			<tr style="height:20">
				<td align=left valign="middle" colspan="3">
					<b>AREA:</b> <%=l_rs("evatitdesabr")%> &nbsp;<br> &nbsp;</td>
				<td valign="middle"  align="right"><%=l_areaResu%>&nbsp;</td>	
				<td><%=l_areaGrabar%> <br>&nbsp;<%=l_areaGrabar2%> <br> &nbsp;</td>
				<td colspan="1">&nbsp;</td>
				<% for i= 1 to Ubound(l_proyectos) 
						Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
						l_sql = "SELECT evadetevldor.evldrnro, evaarea.evatrnro ,evatipresu.evatrdesabr  " 
						l_sql = l_sql & " FROM evacab "
					   	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro  "
					   	l_sql = l_sql & " INNER JOIN evaarea ON evaarea.evldrnro= evadetevldor.evldrnro "
						l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaarea.evatrnro "
						l_sql = l_sql & " WHERE evacab.empleado="& l_ternro &" AND evacab.evaproynro="& l_proyectos(i)  'l_rs("evaproynro") 
						l_sql = l_sql & " AND evadetevldor.evatevnro=" & cevaluador & " AND evaarea.evatitnro = " & l_rs("evatitnro")
						rsOpen l_rs2, cn, l_sql, 0 
						l_areaResu =""
						if not l_rs2.eof then 
							l_areaResu =l_rs2("evatrdesabr")
						else 
							l_areaResu= "Sin evaluar"
						end if
						l_rs2.Close
						set l_rs2=Nothing  %>
					<td>&nbsp;</td>
					<td><%=l_areaResu%>&nbsp;</td>
		 		<% next %>
			</tr>
<%			l_evatitdesabr = l_rs("evatitdesabr")
		end if 
		
		datosCompetenc l_compResu, l_interpretaciones,l_compDescrip, l_compGrabar, l_rs("evafacnro"), l_rs("evatevnro"), l_rs("evatrnro")   'buscar evaresu
		
		if l_rs("evatevnro") = cint(caconsejado) then  %>
		<tr>
			<td valign="top" align="right"><%=l_rs("evafacdesabr")%></td>	
			<td nowrap valign="top"  align="right"> <%=l_compResu%> <%=l_interpretaciones%> </td>
			<td nowrap valign="top"><%=l_compGrabar%></td>
		<% else %>
			<td nowrap valign="top"  align="right"> <%=l_compResu%><%=l_interpretaciones%>	</td>  		
			<td nowrap valign="top"><%=l_compGrabar%></td>
			<% datosObservables l_observables %>
			<td valign=top align=center><%= l_observables%> &nbsp;</td>
		
			<% 	' -------------------------------------------------------------------
		   		'   buscar resultados de la RDE para cada proyecto RDE     
		   		' -------------------------------------------------------------------
		 	For i= 1 to Ubound(l_proyectos) 
				
				Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT distinct evadetevldor.evldrnro, evaresultado.evatrnro ,evatipresu.evatrdesabr,evaobliorden  " 
				l_sql = l_sql & " FROM evacab "
			   	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro  "
			   	l_sql = l_sql & " INNER JOIN evaresultado ON evaresultado.evldrnro= evadetevldor.evldrnro "
				l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
				l_sql = l_sql & " INNER JOIN evaoblieva    ON evaoblieva.evaseccnro=evaseccfactor.evaseccnro AND evadetevldor.evatevnro=evaoblieva.evatevnro"
				l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
				l_sql = l_sql & " WHERE evacab.empleado="& l_ternro &" AND evacab.evaproynro="& l_proyectos(i)  'l_rs("evaproynro") 
				l_sql = l_sql & " AND  (evadetevldor.evatevnro=" & cevaluador & " OR evadetevldor.evatevnro=" & cautoevaluador & ")"
				l_sql = l_sql & " AND evaresultado.evafacnro = " & l_rs("evafacnro")
				l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
				rsOpen l_rs2, cn, l_sql, 0 
				if not l_rs2.eof then 
				 	response.write "<td>"& l_rs2("evatrdesabr")&"</td> " 
					l_rs2.MoveNext
					if not l_rs2.eof then
						response.write "<td>"& l_rs2("evatrdesabr")&"</td> "
					else
						response.write "<td>Sin Evaluar</td> "
					end if
				else 
				 	response.write "<td>Sin Evaluar</td> <td>Sin Evaluar</td>"
				end if
			Next %>
		</tr>
		<%	end if 
		l_rs.Movenext 
  loop
l_rs.Close %>

	<!--
   <tr height="20">
	 	<td colspan="7" align="center"> No existen proyectos asociados a el per&iacute;odo o RDE's cerradas.</td>
   </tr>
   </table>
   -->
</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0"> <!-- -->
</iframe>

<iframe name="terminarsecc" src="termsecc_areasyresultados_eva_00.asp?evacabnro=<%=l_evacabnro%>&evaseccnro=<%=l_evaseccnro%>&evldrnro=<%=l_evldrnro%>&evatevnro=<%=l_evatevnro%>" style="visibility:hidden;width:0;height:0">
</iframe>


</body>
</html>
<% cn.close %>