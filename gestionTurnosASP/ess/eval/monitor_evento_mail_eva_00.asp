<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'===========================================================================================
'Archivo	: monitor_evento_mail_eva_00.asp
'Descripción: cargar body y asunto para mandar el email a todos los de la lista
'Autor		: CCRossi
'Fecha		: 26-07-2004
'Modificar	: 24-11-2004-CCRossi- Controlar caracteres raros
'===========================================================================================

'--------------------------------------------------------------------------------------------

'Parametro de entrada
 Dim l_listamail

'Locales
 Dim l_sql
 Dim l_rs

 Dim l_empemail 
 Dim l_usrname 
 Dim l_usrmail 
 Dim l_nombre
 Dim l_nombre2
 
l_listamail = request("listamail")
l_usrname	= "RRHH"
l_usrmail	= "RRHH"

if trim(l_listamail)="" then%>
	<script>
		window.close();
	</script>
<%end if%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Monitor - Env&iacute;o de mail - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
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

function Validar_Formulario()
{
	document.datos.Subject.value = Blanquear(document.datos.Subject.value);
	document.datos.Body.value    = Blanquear(document.datos.Body.value);
	
	if (document.datos.Subject.value.trim() == "")
		alert('Ingrese un asunto para el email.')
	else
	if (document.datos.Body.value.trim() == "")
		alert('Ingrese un cuerpo para el email.')
	else
		document.datos.submit();
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<table border="0" cellspacing="0" cellpadding="0">
	<tr style="border-color :CadetBlue;">
	<td align="left" class="barra">Monitor - Env&iacute;o de Mail</td>
	<td align="right" class="barra"><a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a></td>
	</tr>
</table>


<table border="0" cellspacing="0" cellpadding="0">
<form name="datos" method="POST" action="monitor_evento_mail_eva_01.asp">
<input type="Hidden" name="listamail"	value="<%= l_listamail%>"> 
<tr>
	<td valign=center><b>Asunto:</b></td>
	<td valign=center>
		<input type="text" name="Subject" size="80">
	</td>
</tr>
<tr>		
	<td valign=top><b>Cuerpo del mail:</b></td>
	<td>
		<textarea name="Body"  maxlength=400 size=400 cols=50 rows=8></textarea>
	</td>
</tr>
</form>
</table>

<table border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</body>
</html>
