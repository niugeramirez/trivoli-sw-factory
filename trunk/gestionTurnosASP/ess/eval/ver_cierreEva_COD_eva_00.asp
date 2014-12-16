<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_cierreEva_COD_eva_00.asp
'Objetivo : 3 evaluacion
'Fecha	  : 16-02-2005
'Autor	  : CCRossi
'Modificacion: 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

' Variables
' de uso local  
  Dim l_existe  
  Dim l_evareunion
  dim l_evaacuerdo
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa

  dim l_nohaydatos  
  dim l_rs2
  dim l_notafinal  
    
  Dim l_evacabnro
  Dim l_evatevnro
  Dim l_lista
  Dim l_primero
  
  dim l_caracteristica  
  dim l_nombre
  
       
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_evaseccnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evaetapa = 3
  
'buscar la evacab
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
	l_evacabnro = l_rs1("evacabnro")
	l_evatevnro = l_rs1("evatevnro")
 end if
 l_rs1.close
 set l_rs1=nothing

' Crear registros de evacierre 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")	
 l_sql = "SELECT DISTINCT  evadetevldor.evldrnro "
 l_sql = l_sql & " FROM evadetevldor "
 l_sql = l_sql & " WHERE evadetevldor.evacabnro  = " & l_evacabnro
 l_sql = l_sql & "   AND evadetevldor.evaseccnro = " & l_evaseccnro
 rsOpen l_rs, cn, l_sql, 0 
 l_lista="0"
 do until l_rs.eof
   l_lista = l_lista & "," & l_rs("evldrnro")
   
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
 	l_sql = l_sql & " FROM  evacierre "
 	l_sql = l_sql & " WHERE evacierre.evldrnro   = " & l_rs("evldrnro")
 	l_sql = l_sql & "   AND evacierre.evaetapa   = " & l_evaetapa
	rsOpen l_rs1, cn, l_sql, 0
	'response.write(l_sql)
	if l_rs1.EOF then
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "insert into evacierre "
		l_sql = l_sql & "(evldrnro,evaetapa,evaacuerdo) "
		l_sql = l_sql & "values (" & l_rs("evldrnro") &","&l_evaetapa &",-1)"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs1.Close
	set l_rs1=nothing
	
   l_rs.MoveNext
 loop 
 l_rs.Close
 set l_rs=nothing 

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaacuerdo "
l_sql = l_sql & " FROM  evacierre "
l_sql = l_sql & " WHERE evacierre.evldrnro =" & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0
if not l_rs1.eof then
	l_evaacuerdo=l_rs1("evaacuerdo")
else
	l_evaacuerdo=-1
end if
l_rs1.Close
set l_rs1=nothing   
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cierre de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
String.prototype.trim = function() {

 // skip leading and trailing whitespace
 // and return everything in between
  var x=this;
  x=x.replace(/^\s*(.*)/, "$1");
  x=x.replace(/(.*?)\s*$/, "$1");
  return x;
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


function Controlar (texto)
{
	texto.value=Blanquear(texto.value);
	
		
	if (texto.value.trim()==""){
		alert('Ingrese un Comentario.');
		texto.focus();
		return false;
	}
	else
		return true;
}	

function garante(ver)
{
	if (ver=='0')
	 document.ifrmgarante.location = 'ver_cierreEva_COD_eva_01.asp?evldrnro=<%=l_evldrnro%>&evaseccnro=<%=l_evaseccnro%>&mostrar=1';
	else
	 document.ifrmgarante.location = 'ver_cierreEva_COD_eva_01.asp?evldrnro=<%=l_evldrnro%>&evaseccnro=<%=l_evaseccnro%>&mostrar=0';		
}
</script>

<style>
.ifrm
{
	font-size: 10;
	border-style: none;
	BACKGROUND-COLOR: #faf0e6;
	scrollbars: no;
}
.rev
{
	font-size: 11;
	border-style: none;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="garante(<%=l_evaacuerdo%>);">
<form name="datos">
<input type="hidden" name="evafecha" value="<%=Date%>">

<table border="0" cellpadding="0" cellspacing="0" height="97%">

<%
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacierre.evldrnro, evareunion, evaacuerdo, evadetevldor.evatevnro, evatevdesabr , evaobser, "
l_sql = l_sql & " empleado.empleg"
l_sql = l_sql & " FROM  evacierre "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro"
l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
l_sql = l_sql & " WHERE evacierre.evldrnro IN (" & l_lista & ")"
rsOpen l_rs1, cn, l_sql, 0
l_primero=-1
l_nohaydatos=0
if l_rs1.eof then
	l_nohaydatos=1
end if

do while not l_rs1.eof
    
    l_caracteristica = "readonly style='background : #e0e0de;'"
	if Int(l_evldrnro) <> l_rs1("evldrnro") then
		l_nombre = l_evldrnro
	else	
		l_nombre = ""
	end if
	if l_primero=-1 then%>
		<tr>
		 	<td width="50%" align=right><b>Cierre y Aprobaci&oacute;n con:</b></td>
		 	<td align=left><%if l_rs1("evaacuerdo")=0 then%>DESACUERDO<%else%>ACUERDO<%end if%></td>
		</tr>
			
		<tr>
		 	<td width="50%" align=right>
		 		<b>¿Se realiz&oacute; la reuni&oacute;n de Evaluaci&oacute;n Final?</b>
		 	</td>
		 	<td align=left><%if l_rs1("evareunion")=0 then%>NO<%else%>SI<%end if%></td>
		</tr>
	<%
	l_primero=0
	end if 
	
	if l_rs1("evatevnro")= cautoevaluador OR l_rs1("evatevnro")=cevaluador then%>
	<tr>		
		<td width="50%" align=right>
			<b>Comentarios <%=l_rs1("evatevdesabr")%>:</b>
		</td>
		<td align=left>
			<textarea <%=l_caracteristica%> name="evaobser<%=l_rs1("evldrnro")%>" maxlength=200 size=200 cols=70 rows=3><%=trim(l_rs1("evaobser"))%></textarea>
		</td>
	</tr>
	
	<%end if
	
	l_rs1.Movenext
	
Loop
l_rs1.close
set l_rs1=nothing

if l_nohaydatos=0 then

	'Buscar resultados de cada ambito
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evatipobjdabr, evapuntaje.puntaje, evatipopor "
	l_sql = l_sql & " FROM  evapuntaje"
	l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro=evapuntaje.evacabnro"
	l_sql = l_sql & " INNER JOIN evatipoobjpor ON evatipoobjpor.evatipobjnro=evapuntaje.evatipobjnro"
	l_sql = l_sql & "		 AND evatipoobjpor.evaevenro=evacab.evaevenro " 
	l_sql = l_sql & " INNER JOIN evatipoobj    ON evatipoobjpor.evatipobjnro=evatipoobj.evatipobjnro"
	l_sql = l_sql & " WHERE evapuntaje.evacabnro = " & l_evacabnro
	rsOpen l_rs2, cn, l_sql, 0
	'Response.Write l_sql
	l_notafinal= 0
	do while not l_rs2.EOF %>
		<tr>		
			<td align=right width="50%">
				<b><%=l_rs2("evatipobjdabr")%>&nbsp;(<%=l_rs2("evatipopor")%>%):</b>
			</td>
			<td align=left>
				<b>&nbsp;<%=l_rs2("puntaje")%></b>
			</td>
		</tr>
		<%	
		l_notafinal= l_notafinal + cdbl(l_rs2("puntaje")) * cdbl(l_rs2("evatipopor")) / 100
		l_rs2.Movenext
	loop
	l_rs2.close
	set l_rs2=nothing
	%>
	<tr>		
		<td align=right width="50%"><b>NOTA FINAL PROPUESTA:</b></td>
		<td align=left><input readonly style='background : #e0e0de;' type="text" name="notafinalpropuesta" size=5 value="<%=l_notafinal%>"></td>
	</tr>
	<tr height="30%">		
		<td align=right colspan=2>
		  <iframe src="blanc.asp" name="ifrmgarante" class="ifrm" scrolling="No" frameborder="0" width="100%" height="100" ></iframe>
		</td>
	</tr>

		
<%end if
cn.Close
set cn = Nothing
%>	
<tr>
	<td valign=top align=right colspan=2>
	</td>
</tr>
</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0"></iframe>
</body>
</html>
