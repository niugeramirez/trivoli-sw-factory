<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: carga_calificompSI_eva_00.asp
'Descripción	: Cargar competencias y areas para varios evaluadores
'Autor			: 29-08-2005
'Fecha			: Leticia Amadio
'Modificado		: Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion
'				  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================
on error goto 0

' Variables
 
' de uso local  
  dim l_evatrnro 
  dim l_evacabnro 
  dim l_ternro  
  dim l_evaluador 
  dim l_evaevenro
  dim l_datos 

 ' dim l_horas
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

 ' dim l_areatrnro
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

dim l_evaluadores (10)
dim l_tipoevaluadores (10)
dim l_evldrnros(10)
dim l_cantevldores 

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  Dim l_evapernro  ' VERRR si se pasa siempre este parametro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")
  l_evapernro  = Request.QueryString("evapernro")
  
'_______________________________________________________________________________
' fijarse que halla posibles resultados configurado, para la seccion,
'  los posibles resultados se definen para cada competencia
 
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
end if  
l_rs1.Close
set l_rs1 = nothing 

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


'buscar las estructuras ACTIVAS del empleado --------------------------------------------------
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


' ________________________________________________________
' ________________________________________________________
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT  evatipevalua.evatevdesabr, evatipevalua.evatevnro, evadetevldor.evldrnro, evaoblieva.evaobliorden "
l_sql = l_sql & " FROM  evatipevalua "
l_sql = l_sql & " INNER JOIN  evaoblieva ON evaoblieva.evatevnro = evatipevalua.evatevnro " 
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evatevnro = evaoblieva.evatevnro AND evadetevldor.evaseccnro = evaoblieva.evaseccnro"
l_sql = l_sql & " WHERE  evaoblieva.evaseccnro ="& l_evaseccnro & " AND evadetevldor.evacabnro ="& l_evacabnro
l_sql = l_sql & " ORDER BY evaobliorden "
rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
i = 1
do while not l_rs.eof 
	l_evaluadores(i)= l_rs("evatevdesabr")
	l_tipoevaluadores(i)=l_rs("evatevnro")
	l_evldrnros(i) = l_rs("evldrnro")
	i= i+1
l_rs.MoveNext
loop
l_cantevldores = i-1
l_rs.Close 



' __________________________________________________________________________________________
' Crear registros de evaresultado para los facnro y el evldrnro (para evaluado y evaluador) 
' Crear reg para las areas -evaarea 														
' __________________________________________________________________________________________
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT DISTINCT evaseccfactor.evafacnro, evatitulo.evatitnro, evaseccfactor.evaseccnro, evatitulo.evatitdesabr " ' evaresu.evatrnro, 
l_sql = l_sql & " FROM evaseccfactor " 
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " WHERE evaseccfactor.evaseccnro ="& l_evaseccnro
l_sql = l_sql & " ORDER BY evatitulo.evatitnro "
rsOpen l_rs, cn, l_sql, 0 

set l_cm = Server.CreateObject("ADODB.Command")

if not l_rs.eof then
  	l_evatitdesabr = ""
	l_evatrnro  = "NULL"
	do while not l_rs.eof
		if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then 
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			For i = 1 to l_cantevldores 
				l_sql = "SELECT *  FROM  evaarea "
				l_sql = l_sql & " WHERE evldrnro="&l_evldrnros(i) &" AND evaarea.evatitnro ="& l_rs("evatitnro")
				rsOpen l_rs1, cn, l_sql, 0
				if l_rs1.EOF and l_tipoevaluadores(i) <> cint(cautoevaluador) then 
					l_sql = "INSERT INTO evaarea  (evldrnro, evatitnro, evatrnro, evaareadesc) "
					l_sql = l_sql & " VALUES ("& l_evldrnros(i) & "," & l_rs("evatitnro")& ","& l_evatrnro & ",'')"
					l_cm.activeconnection = Cn
					l_cm.CommandText = l_sql
					cmExecute l_cm, l_sql, 0
				end if
				l_rs1.Close
			Next
			set l_rs1=nothing 
			l_evatitdesabr = l_rs("evatitdesabr") 
		end if 
	  	
  		Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
		For i = 1 to l_cantevldores 
			l_sql = "SELECT *  FROM  evaresultado "
			l_sql = l_sql & " WHERE evldrnro="&l_evldrnros(i) &" AND evafacnro ="& l_rs("evafacnro")
			rsOpen l_rs1, cn, l_sql, 0
			if l_rs1.EOF then 
				l_sql = "INSERT INTO evaresultado  (evldrnro, evafacnro, evatrnro, evaresudesc) "
				l_sql = l_sql & " VALUES (" & l_evldrnros(i) & "," & l_rs("evafacnro") & ","& l_evatrnro & ",'')"
				l_cm.activeconnection = Cn 
				l_cm.CommandText = l_sql   
				cmExecute l_cm, l_sql, 0   
			end if
			l_rs1.Close
		Next
		set l_rs1=nothing 
	
	l_rs.MoveNext 
	loop 
end if 
l_rs.Close
set l_rs=nothing
' -----------------------------------------------------------------------------------


' _____________________________________________________________________
'  		
' ______________________________________________________________________
sub datosArea (evatitnro, evldrnro, areaResu,  areaGrabar, areaGrabar2) 
dim l_areatrnro
areaResu=""

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evatrnro FROM evaarea "
l_sql = l_sql & " WHERE evaarea.evatitnro ="& evatitnro & " AND evldrnro="& evldrnro 
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then 
	l_areatrnro=l_rs1("evatrnro") 
end if 
l_rs1.close 
set l_rs1=nothing 

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaresu.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr  "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaresu.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro ="& l_evaseccnro & " AND   evatitulo.evatitnro="& evatitnro  'l_rs("evatitnro")
l_sql = l_sql & " ORDER BY evatrvalor "
rsOpen l_rs1, cn, l_sql, 0 

if cint(l_evldrnro) <> cint(evldrnro)   then 'l_rs("evatitnro")
	areaResu= areaResu & "<select name=areatrnro"&evatitnro&evldrnro&" disabled style="" width:100 "">"
else 
	areaResu= areaResu & "<select name=areatrnro"&evatitnro&evldrnro&" style="" width:100 "">"
end if

areaResu= areaResu & "<option value=>&nbsp;Sin Evaluar</option> "
do while not l_rs1.eof
	areaResu= areaResu & "<option value="&l_rs1("evatrnro")&">&nbsp;&nbsp;"& l_rs1("evatrdesabr")&"</option> "
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
areaResu= areaResu & "	</select> "
			'l_rs("evatitnro")
areaResu= areaResu & "  <script>document.datos.areatrnro"&evatitnro&evldrnro&".value='"&l_areatrnro&"';</script>"

if cint(l_evldrnro) <> cint(evldrnro) then 
	areaGrabar="<a href=# onclick=""return false;"">Grabar</a>"	
else 
    areaGrabar= "<a href=# onclick=""if (Controlar(document.datos.areatrnro"&evatitnro&evldrnro&")) {grabar.location='grabar_areas_evaluacion_00.asp?evatitnro="&evatitnro&"&evldrnro="&evldrnro&"&evatrnro='+document.datos.areatrnro"&evatitnro&evldrnro&".value+'&campo=areatrnro"&evatitnro&evldrnro&"';document.datos.grabado"&evatitnro&evldrnro&".value='G'; }"">Grabar</a>"
						' &evaareadesc='+escape(Blanquear(document.datos.evaareadesc"&l_rs("evatitnro")&".value))+'
end if   
areaGrabar2= "<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabado"&evatitnro&evldrnro&"""  size=""1"">"

end sub   


' ____________________________________________________________________________________________
' ____________________________________________________________________________________________
sub datosCompetenc (evafacnro, evldrnro, compResu,interpr, compGrabar) 'buscar evaresu
dim l_comptrnro
compResu = ""			' , evatevnro, evatrnro, compDescrip
'compDescrip = ""
interpr=""


' Interpretaciones 										
' OBS -- > se hace en gral y no por rol y la competencia - El Estandar NO carga la descripciones cuando se copian Resu!
Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT distinct evaresu.evatrnro,evaresu.evaresudes, evatipresu.evatrvalor, evatipresu.evatrdesabr " '
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro 
l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rs("evafacnro") & " AND NOT (evaresudes Is Null) "
l_sql = l_sql & " ORDER BY evaresu.evatrnro "
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
l_sql = "SELECT evatrnro FROM evaresultado "
l_sql = l_sql & " WHERE evaresultado.evafacnro ="& evafacnro & " AND evldrnro="& evldrnro 
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then 
	l_comptrnro=l_rs1("evatrnro") 
end if 
l_rs1.close 
set l_rs1=nothing 


Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT  DISTINCT evaresu.evatrnro, evaresu.evaresudes, evatipresu.evatrvalor, evatipresu.evatrdesabr"
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro="& l_evaseccnro & " AND  evaresu.evafacnro="& evafacnro
l_sql = l_sql & " ORDER BY evaresu.evatrnro "
rsOpen l_rs1, cn, l_sql, 0 

if cint(l_evldrnro) <> cint(evldrnro) then
	compResu = compResu & "<select name=evatrnro"&evafacnro&evldrnro&" style="" width:90 "" disabled>"
else
	compResu = compResu & "<select name=evatrnro"&evafacnro&evldrnro&" style="" width:90 "">"
end if
compResu = compResu & "<option value=>Sin Evaluar</option>"
do while not l_rs1.eof
	compResu = compResu & "<option value="& l_rs1("evatrnro")&">"& "&nbsp;"&l_rs1("evatrdesabr")&" </option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
compResu = compResu & "</select> "

compResu = compResu & " <script>document.datos.evatrnro"&evafacnro&evldrnro&".value='"&l_comptrnro&"'</script>"
'else
	'compResu = compResu & " <script>document.datos.evatrnro"&evafacnro&evatevnro&".value=''</script>"
'end if
'compDescrip = "	<input type=hidden name=evaresudesc"&evafacnro&evldrnro&">"
'compDescrip = "<textarea name=evaresudesc"&l_rs("evafacnro")&" cols=30 rows=5>"& trim(l_rs("evaresudesc"))&"</textarea>"


if  cint(l_evldrnro) <> cint(evldrnro) then ' l_cantproysaprob <= 0 or
	compGrabar= "<a href=# >Grabar</a>"
else 																																' "&evaresudesc='+escape(Blanquear(document.datos.evaresudesc"&evafacnro&evatevnro&".value))+'
	compGrabar = "<a href=# onclick=""if (Controlar(document.datos.evatrnro"&evafacnro&evldrnro&")){ grabar.location='grabar_competencias_evaluacion_00.asp?evafacnro="&evafacnro&"&evldrnro="&l_evldrnro&"&evatrnro='+document.datos.evatrnro"&evafacnro&evldrnro&".value+'&campo=evatrnro"&evafacnro&evldrnro&"';document.datos.grabado"&evafacnro&evldrnro&".value='G'; }"">Grabar</a>"
end if																																																																																									' CAMPO=GRABAR ????
compGrabar = compGrabar & "<br>" &"<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabado"&evafacnro&evldrnro&""" size=""1"">"		


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
'l_sql = l_sql & " AND   evadescomp.tenro    =  " & ctenro
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
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
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
<input type="Hidden" name="terminarsecc" value="--">
<input type="Hidden" name="terminarsecc2" value="">

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr height="20">
	<td colspan="<%= 2 + l_cantevldores * 2%>" align="center">
		<b>CALIFICACI&Oacute;N POR COMPETENCIAS </b>
	</td>
</tr>
<tr><td colspan="<%=2 + l_cantevldores * 2 %>">&nbsp;</td> </tr>
<tr>
	<th nowrap class="th2">AREAS/COMPETENCIAS </th>
	<% For i = 1 to l_cantevldores %>
		<th colspan="2" nowrap class="th2"><%= UCase(l_evaluadores(i))%>&nbsp;</th> 
	<% Next	%>
	<th class="th2">OBS. </th>
</tr>
<tr><td colspan="<%= 2 +  l_cantevldores * 2 %>">&nbsp;</td></tr>

<% '
'xxxxxxxxxxxxxxxxxxxxxxxxx
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaseccfactor.evafacnro,evafactor.evafacdesabr, evafactor.evafacdesext, evaseccfactor.orden,evatitulo.evatitnro, evatitulo.evatitdesabr " 
 l_sql = l_sql & " FROM evaseccfactor "
 l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaseccfactor.evafacnro "
 l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
 l_sql = l_sql & " WHERE evaseccfactor.evaseccnro ="& l_evaseccnro
 l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evaseccfactor.orden "  
 rsOpen l_rs, cn, l_sql, 0

 l_evatitdesabr="" 

 do while not l_rs.eof 
	   	' para cada area mostrar su .. evaarea..
	if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then %>
		<tr style="height:20">
			<td align=left valign="middle" colspan="3"><b>AREA:</b> <%=l_rs("evatitdesabr")%> &nbsp;<br> &nbsp;</td>	
		<% For i= 1 to l_cantevldores 
			if cint(l_tipoevaluadores(i)) <> cint(cautoevaluador) then 
				datosArea l_rs("evatitnro"), l_evldrnros(i), l_areaResu, l_areaGrabar, l_areaGrabar2  %>
				<td valign="middle"  align="right"><%=l_areaResu %>&nbsp;</td>
				<td><%=l_areaGrabar%>&nbsp;<br><%=l_areaGrabar2%> &nbsp;</td>
		<%	end if
		 Next %>
			<td colspan="1">&nbsp;</td>
		</tr>
<%			l_evatitdesabr = l_rs("evatitdesabr") 
	end if %>
	
	<tr><td valign="top"><%=l_rs("evafacdesabr")%></td>
	<%  For i= 1 to l_cantevldores	
		datosCompetenc l_rs("evafacnro"), l_evldrnros(i), l_compResu, l_interpretaciones, l_compGrabar %>
		<td nowrap valign="top" align="right"> <%=l_compResu%> <%=l_interpretaciones%> </td>
		<td nowrap valign="top"><%=l_compGrabar%></td>
	<% Next %>
	<% datosObservables l_observables %>
		<td valign=top align=center><%= l_observables%> &nbsp;</td>
	</tr>			
	<% 
	l_rs.Movenext 
loop
l_rs.Close %>

<!--
   <tr height="20">
	 	<td colspan="7" align="center"> No existen proyectos asociados a el per&iacute;odo o RDE's cerradas.</td>
   </tr>
   </table>
-->
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<iframe src="blanc.asp" name="terminarsecc" src="termsecc_areasyresultadosSI_eva_00.asp?evacabnro=<%=l_evacabnro%>&evaseccnro=<%=l_evaseccnro%>&evldrnro=<%=l_evldrnro%>&evatevnro=<%=l_evatevnro%>" style="visibility:hidden;width:0;height:0">
</iframe>
</form>	
</body>
</html>
<% cn.close %>