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
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'================================================================================


on error goto 0

' Variables
 
' de uso local  
  dim l_evatrnro 
  dim l_evacabnro 
  dim l_ternro  
  dim l_evaluador 
  
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

'  dim l_compGrabar2
  Dim l_habCalifArea
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

  ' busca ternro del Evaluado (Aconsejado) y el evacabnro
  l_ternro=""
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evacab.evacabnro, empleado "
  l_sql = l_sql & " FROM evacab "
  l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
  l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then
	l_evacabnro  = l_rs("evacabnro")
	l_ternro   = l_rs("empleado")
  end if
  l_rs.close
  set l_rs=nothing


' ______________________________________________________________
'  Busco si existe un reg de esta evaluacion en evagralobj 
'l_evatrnro  = ""
'Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
'l_sql = "SELECT *  FROM  evagralobj"
'l_sql = l_sql & " WHERE evagralobj.evldrnro="& l_evldrnro
'rsOpen l_rs1, cn, l_sql, 0 
'if l_rs1.EOF then
	'set l_cm = Server.CreateObject("ADODB.Command")  
	'l_sql = "INSERT INTO evagralobj "
	'l_sql = l_sql & " (evldrnro, evatrnro) "
	'l_sql = l_sql & " VALUES (" & l_evldrnro & ",NULL)" 
	'l_cm.activeconnection = Cn
	'l_cm.CommandText = l_sql
	'cmExecute l_cm, l_sql, 0
'else
	'l_evatrnro= l_rs1("evatrnro")
'end if
'l_rs1.Close
'set l_rs1=nothing




Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT evacabnro, evatevnro  " 
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro= " & l_evldrnro 
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then
	l_evacabnro = l_rs1("evacabnro")
	l_evatevnro = l_rs1("evatevnro")
end if  
l_rs1.Close
set l_rs1 = nothing

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

'  response.write l_sql

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




'________________________________________________________________________________
' Si el revisor evaluo todos los objetivos --> se habilita la Calif gral de obj  
' _______________________________________________________________________________
l_habCalifArea="NO"
	
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evadetevldor.evldrnro, evatevnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN evaresultado ON evaresultado.evldrnro = evadetevldor.evldrnro "
 l_sql = l_sql & " WHERE evadetevldor.evacabnro ="& l_evacabnro 
 l_sql = l_sql & "   AND evadetevldor.evatevnro ="& cconsejero  
 l_sql = l_sql & "   AND evadetevldor.evaseccnro="& l_evaseccnro
 l_sql = l_sql & "   AND evaresultado.evatrnro IS NULL " 
  	 'l_sql = l_sql & " WHERE evadetevldor.evldrnro="& l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
'response.write l_sql
 if l_rs1.EOF then
   l_habCalifArea="SI"
 end if 
 l_rs1.close
 set l_rs1=nothing
  
 
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

'response.write l_areatrnro
'response.write l_sql

if l_habCalifArea="NO" then
	areaResu= areaResu & "<select name=areatrnro"& l_rs("evatitnro")&" disabled style="" width:180 "">"
else 
	areaResu= areaResu & "<select name=areatrnro"& l_rs("evatitnro")&" disabled style="" width:180 "">"
end if
	
areaResu= areaResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
do while not l_rs1.eof
	areaResu= areaResu & "<option value="&l_rs1("evatrnro")&">&nbsp;&nbsp;"& l_rs1("evatrdesabr")&"</option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
areaResu= areaResu & "	</select> "
areaResu= areaResu & "  <script>document.datos.areatrnro"&l_rs("evatitnro")&".value='"&l_areatrnro&"';</script>"

'areaDescrip = "<textarea name=evaareadesc"&l_rs("evatitnro")&" cols=30 rows=4 disabled>" & trim(l_evaareadesc)& "</textarea>"

if l_habCalifArea="NO" then
	areaGrabar="<a href=# onclick=""return false;"">Grabar</a>"	
else 
    areaGrabar= "<a href=# onclick=""if (Controlar(document.datos.areatrnro"&l_rs("evatitnro")&")) {grabar.location='grabar_areas_evaluacion_00.asp?evatitnro="&l_rs("evatitnro")&"&evldrnro="&l_evldrnro&"&evatrnro='+document.datos.areatrnro"&l_rs("evatitnro")&".value+'&campo=areatrnro"&l_rs("evatitnro")&"';document.datos.grabado"&l_rs("evatitnro")&".value='G'; }"">Grabar</a>"
						' &evaareadesc='+escape(Blanquear(document.datos.evaareadesc"&l_rs("evatitnro")&".value))+'
end if   
areaGrabar2= "<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabado"&l_rs("evatitnro")&"""  size=""1"">"

end sub   

' ____________________________________________________________________________
' ____________________________________________________________________________
sub datosCompetenc (compResu,interpr, compDescrip, compGrabar, evafacnro, evatevnro, evatrnro) 'buscar evaresu
compResu = ""
compDescrip = ""
interpr=""

'response.write evafacnro & " - " & evatevnro & " - " & evatrnro & "<br>"
		
		' ARREGLARRRRRRRRRRRRRRRRRRRRRR -- > ver si se puede cambiar segun el rol? y la competencia???
Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT distinct evaresu.evatrnro,evaresu.evaresudes,  " '
l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
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
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro 
l_sql = l_sql & " AND   evaresu.evafacnro  = " & evafacnro
l_sql = l_sql & " order by evaresu.evatrnro "
rsOpen l_rs1, cn, l_sql, 0 

'response.write l_sql & "<br>"
if l_evatevnro <> evatevnro then
	compResu = compResu & "<select name=evatrnro"&evafacnro&evatevnro&" style="" width:150 "" disabled>"
else
	compResu = compResu & "<select name=evatrnro"&evafacnro&evatevnro&" style="" width:150 "" disabled>"
end if
compResu = compResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
do while not l_rs1.eof
	'if trim(l_rs1("evaresudes")) <>"" and not isnull(l_rs1("evaresudes")) then
		'interpr = interpr & l_rs1("evatrdesabr") &": "&unescape(l_rs1("evaresudes")) & "\n"
	'end if
	compResu = compResu & "<option value="& l_rs1("evatrnro")&">"& "&nbsp;"&l_rs1("evatrdesabr")&" </option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
compResu = compResu & "</select> "

'if evatrnro <> "" then
	compResu = compResu & " <script>document.datos.evatrnro"&evafacnro&evatevnro&".value='"& evatrnro&"'</script>"
'else
	'compResu = compResu & " <script>document.datos.evatrnro"&evafacnro&evatevnro&".value=''</script>"
'end if
compDescrip = "	<input type=hidden name=evaresudesc"&evafacnro&evatevnro&">"
'compDescrip = "<textarea name=evaresudesc"&l_rs("evafacnro")&" cols=30 rows=5>"& trim(l_rs("evaresudesc"))&"</textarea>"


if l_evatevnro <> evatevnro then
	compGrabar= "<a href=# >Grabar</a>"
else 																																' "&evaresudesc='+escape(Blanquear(document.datos.evaresudesc"&evafacnro&evatevnro&".value))+'
	compGrabar = "<a href=# onclick=""if (Controlar(document.datos.evatrnro"&evafacnro&evatevnro&")){ grabar.location='grabar_competencias_evaluacion_00.asp?evafacnro="&evafacnro&"&evldrnro="&l_evldrnro&"&evatrnro='+document.datos.evatrnro"&evafacnro&evatevnro&".value+'&campo=evatrnro"&evafacnro&evatevnro&"';document.datos.grabado"&evafacnro&evatevnro&".value='G'; }"">Grabar</a>"
end if																																																																																									' CAMPO=GRABAR ????
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
  
  
  
'buscar el ternro del EVALUADO --------------------------------------------------------
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
		texto.focus();
		return false;
	}
	else
		if (valor.value==""){
			alert('Seleccione un resultado.');
			valor.focus();
			return false;
		}
		else
			return true;
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
<input type="Hidden" name="terminarsecc" value="SI">
<input type="Hidden" name="terminarsecc2" value="">

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr height="20">
	<td colspan="7">
	<table><tr>
		<td align="center" width="25%" colspan=4><b>CALIFICACI&Oacute;N POR COMPETENCIAS RDE&rsquo;s</b></td> 
	</tr></table>
	</td>
</tr>
<tr height="20"><td colspan="7" align="center">&nbsp;</td> </tr>
<tr>
	<th>Proyecto </th>
	<th>Engagement </th>
	<th>Cliente </th>
	<th>Gerente </th>
	<th>Socio </th>
	<th>Horas <br> Imputadas </th>
	<th>Desde</th> 
	
</tr>
<%
 	' buscar lista de engagement en la que participo el empleado, en un periodo dado. -----------------
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaengage.evaengnro, evaengdesabr, evaclinom, proygerente, proysocio, evaproyfdd, evaproyecto.evaproynro"
 l_sql = l_sql & " FROM evaengage "
 l_sql = l_sql & " INNER JOIN evacliente  ON evacliente.evaclinro  = evaengage.evaclinro "
 l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evaengnro = evaengage.evaengnro "
 l_sql = l_sql & " INNER JOIN evaproyemp  ON evaproyemp.evaproynro = evaproyecto.evaproynro "
 l_sql = l_sql & " WHERE evaproyemp.ternro = " & l_ternro & " AND evaproyecto.evapernro=" & l_evapernro
 rsOpen l_rs, cn, l_sql, 0 
 'response.write l_sql 
 
 if not l_rs.eof then 
 	
	do while not l_rs.eof 
		
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
%>
		<tr>
			<td><strong><%= l_rs("evaproynro")%></strong></td>
			<td><strong><%= l_rs("evaengdesabr")%></strong></td>
			<td><strong><%= l_rs("evaclinom")%></strong></td>
			<td><strong><%= l_gerente %></strong></td>
			<td><strong><%= l_socio %></strong></td>
			<td><strong><%= l_horas%></strong></td>
			<td><strong><%= l_rs("evaproyfdd")%></strong></td> 
		</tr>
		<tr style="height:25">
			<td colspan="4" align="center"> <strong>AREAS</td>
			<td><strong>COMPETENCIAS </strong></td>
			<td><strong>AUTOEVALUADOR</strong></td>
			<td><strong>REVISOR</strong></td>
		</tr>
<%  			
		' buscar datos del areas y competencias..............
	   l_sql = "SELECT  distinct evaresultado.evldrnro, evaresultado.evafacnro,evaresultado.evatrnro ,evatipresu.evatrdesabr" '
	   l_sql = l_sql & " ,evafactor.evafacdesabr, evafactor.evafacdesext, evaseccfactor.orden "
	   l_sql = l_sql & " ,evatitulo.evatitnro, evatitulo.evatitdesabr, evadetevldor.evatevnro "
	   l_sql = l_sql & " FROM evacab "
	   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro  "
	   l_sql = l_sql & " INNER JOIN evaresultado ON evaresultado.evldrnro= evadetevldor.evldrnro "
	   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
	   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
	   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
	   l_sql = l_sql & " LEFT JOIN evaresu ON evaresu.evatrnro = evaresultado.evatrnro and evaresu.evafacnro =evaseccfactor.evafacnro"
	   l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
	   l_sql = l_sql & " WHERE evacab.empleado =" & l_ternro & " AND evacab.evaproynro=" & l_rs("evaproynro")
	   l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evaseccfactor.orden, evadetevldor.evatevnro" 

	   '	response.write l_sql
		rsOpen l_rs1, cn, l_sql, 0 
		
		   l_evatitdesabr="" 
			
		   do while not l_rs1.eof 
				if trim(l_evatitdesabr) <> trim(l_rs1("evatitdesabr")) then
					Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT evadetevldor.evldrnro, evaarea.evatrnro ,evatipresu.evatrdesabr  " 
					l_sql = l_sql & " FROM evacab "
				   	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro  "
				   	l_sql = l_sql & " INNER JOIN evaarea ON evaarea.evldrnro= evadetevldor.evldrnro "
					l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaarea.evatrnro "
					l_sql = l_sql & " WHERE evacab.empleado="& l_ternro &" AND evacab.evaproynro="& l_rs("evaproynro") 
					l_sql = l_sql & " AND evadetevldor.evatevnro=" & cevaluador & " AND evaarea.evatitnro = " & l_rs1("evatitnro")
					rsOpen l_rs2, cn, l_sql, 0 
					l_areaResu =""
					if not l_rs2.eof then
						l_areaResu =l_rs2("evatrdesabr")
					end if
					l_rs2.Close
					set l_rs2=Nothing
		%>
					<tr style="height:20">
						<td align=right colspan="4" valign="bottom">
							AREA: <%=l_rs1("evatitdesabr")%> &nbsp;   &nbsp; &nbsp;		
						</td>
						<td><%= l_areaResu %>&nbsp;</td>
						<td colspan="2"> &nbsp; &nbsp;</td>
					</tr>
		<%			l_evatitdesabr = l_rs1("evatitdesabr")
				end if 
				
				'datosCompetenc l_compResu, l_interpretaciones,l_compDescrip, l_compGrabar, l_rs1("evafacnro"), l_rs1("evatevnro"), l_rs1("evatrnro")   'buscar evaresu
		
				if l_rs1("evatevnro") = cint(cautoevaluador) then 
		%>
				<tr>
					<td>&nbsp;</td>
					<td valign="top" colspan="4" align="right"><%=l_rs1("evafacdesabr")%></td>	
					<td nowrap valign="top"> <%=l_rs1("evatrdesabr")%>&nbsp;&nbsp;</td>
		<%   	else %>	
					<td nowrap valign="top"><%=l_rs1("evatrdesabr")%>		</td>
				</tr>
		<% 
			   end if 
		l_rs1.Movenext 
		loop
		%>
			<tr><td colspan=7>&nbsp;</td></tr>
		<%
		l_rs1.Close
		
		
	l_rs.MoveNext
	loop 

else  %>
   <tr height="20">
	 	<td colspan="7" align="center"> No existen proyectos asociados a el per&iacute;odo.</td>
   </tr>
<%
end if
  
l_rs.close
set l_rs=nothing
%>
</table>

<% ' XXXXXXXXXXXXXXXXXXXXXXX
	'	Formulario para cargar datos..
%>
<table>
<tr>
	<td align="center" width="25%" colspan=8>&nbsp;</td> 
	</tr>
<tr>
	<td align="center" width="25%" colspan=8><b>CALIFICACI&Oacute;N POR COMPETENCIAS A CARGAR</b></td> 
</tr>
<tr height="15" class="th2">
   	<th align=center colspan=3 class="th2">&nbsp; &nbsp;&nbsp;&nbsp; COMPETENCIAS &nbsp;&nbsp; &nbsp;&nbsp;</th> <!-- Descripci&oacute;n -->
	
	<th colspan="2" class="th2">AUTOEVALUADOR</th> 
	<th colspan="2" class="th2">REVISOR</th> 
	<th class="th2">OBSERVABLES</th>
</tr>

<%'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaresultado.evldrnro, evaresultado.evafacnro,evaresultado.evatrnro " 'evaresultado.evaresudesc
   l_sql = l_sql & " ,evafactor.evafacdesabr, evafactor.evafacdesext, evaseccfactor.orden"
   l_sql = l_sql & " ,evatitulo.evatitnro, evatitulo.evatitdesabr, evadetevldor.evatevnro, evaoblieva.evaobliorden "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " INNER JOIN evaoblieva    ON evaoblieva.evaseccnro=evaseccfactor.evaseccnro "
   'l_sql = l_sql & " INNER JOIN evaresultado  ON evadetevldor.evldrnro = evaresultado.evldrnro "
   l_sql = l_sql & " LEFT JOIN evadetevldor  ON evadetevldor.evldrnro = evaresultado.evldrnro AND evadetevldor.evatevnro=evaoblieva.evatevnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro =" & l_evaseccnro
   'l_sql = l_sql & "   AND evaresultado.evldrnro    =" & l_evldrnro 
   l_sql = l_sql & "   AND evadetevldor.evacabnro="& l_evacabnro
   l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evaseccfactor.orden, evaoblieva.evaobliorden"
   
  'response.write l_sql 
   
   l_evatitdesabr="" 
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof 
		if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then
			datosArea l_areaResu, l_areaDescrip, l_areaGrabar, l_areaGrabar2  'buscar evaarea 
%>
			<tr style="height:25"><td colspan="8">&nbsp;</td></tr>
			<tr style="height:20">
				<td align=right colspan="3"  valign="middle">
					<b>AREA: <%=l_rs("evatitdesabr")%> </b>&nbsp;   <br> &nbsp; &nbsp;

				</td>
				
				<td colspan="5" valign="middle">
					<%=l_areaResu%> &nbsp; &nbsp;	
					<%=l_areaGrabar%> &nbsp;<%=l_areaGrabar2%> <br> &nbsp;
				</td>
			</tr>
<%			l_evatitdesabr = l_rs("evatitdesabr")
		end if 
		
		datosCompetenc l_compResu, l_interpretaciones,l_compDescrip, l_compGrabar, l_rs("evafacnro"), l_rs("evatevnro"), l_rs("evatrnro")   'buscar evaresu

		if l_rs("evatevnro") = cint(caconsejado) then 
		'response.write "entre aconsejado <br>"
%>
		<tr>
			<td valign="top" colspan="2"><%=l_rs("evafacdesabr")%></td>	
			<td>&nbsp;</td>
			<td nowrap valign="top"> 
				<%=l_compResu%> <%=l_interpretaciones%> 	
			</td>
			<td nowrap valign="top">
				<%=l_compGrabar%>
			</td>
<%   	else 
			'response.write "entre consejero <br>"
%>
			<td nowrap valign="top"> &nbsp;
				<%=l_compResu%><%=l_interpretaciones%>	
			</td>  		
			<td nowrap valign="top">&nbsp;
				<%=l_compGrabar%>
			</td>
<% 			 datosObservables l_observables %>
			<td valign=top align=center><%'= l_observables%> &nbsp;</td>
		</tr>
<% 
	   end if 
	l_rs.Movenext 
  loop
l_rs.Close%>
</form>	
</table>

<iframe name="grabar" style="visibility:hidden;width:0;height:0"> <!-- -->
</iframe>
<!--
<iframe name="terminarsecc" src="termsecc_areasyresultados_eva_00.asp?evacabnro=<%=l_evacabnro%>&evaseccnro=<%=l_evaseccnro%>&evldrnro=<%=l_evldrnro%>&evatevnro=<%=l_evatevnro%>&habCalifArea=<%=l_habCalifArea%>" style="visibility:hidden;width:0;height:0">
</iframe>
-->


</body>
</html>
<%
cn.close
%>