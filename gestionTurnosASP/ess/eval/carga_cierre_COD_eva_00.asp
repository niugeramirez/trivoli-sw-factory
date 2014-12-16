<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_cierre_COD_eva_00.asp
'Objetivo : Cierre de una etapa (1: planificacion, 2: seguimiento, 3 evaluacion)
'Fecha	  : 05-02-2005
'Autor	  : CCRossi
'Modificacion: 13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

' Variables
' de uso local  
  Dim l_existe  
  Dim l_evareunion
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  l_evaetapa = 1
  
' Crear registros de evafirm para evldrnro y el tipo nota
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT * FROM  evacierre WHERE evacierre.evldrnro =  " & l_evldrnro & "   AND evacierre.evaetapa =1 "
   rsOpen l_rs1, cn, l_sql, 0
'  response.write(l_sql)
   if l_rs1.EOF then
		l_evareunion=0
  		l_evafecha= Date()
		l_evaobser=""
   else
  		l_evareunion=l_rs1("evareunion")
  		l_evafecha= l_rs1("evafecha")
		l_evaobser= l_rs1("evaobser")
   end if
   l_rs1.Close
   set l_rs1=nothing
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cierre de la Planificaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
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

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

</script>

<style>
.blanc
{
	font-size: 10;
	border-style: none;
	background : transparent;
}
.rev
{
	font-size: 10;
	border-style: none;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="<%if l_evareunion=0 then%>document.datos.evareunion.value=0<%else%>document.datos.evareunion.value=-1<%end if%>;">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
	
<tr>
 	<td align=right>
 		<b>¿Se realiz&oacute; la reuni&oacute;n de Planificaci&oacute;n?</b>
 	</td>
 	<td align=left>
 		<input type="hidden" name="evareunion" value="<%=l_evareunion%>">
 		<input type="Radio" onclick="document.datos.evareunion.value=0;" name="radio" value="0" <%If l_evareunion = 0 then%>checked<%End If%>>NO
 		<input type="Radio" onclick="document.datos.evareunion.value=-1;" name="radio" value="-1" <%If l_evareunion = -1 then%>checked<%End If%>>SI
 	</td>
</tr>
<tr>		
	<td align=right>
		<b>Observaci&oacute;n:</b>
	</td>
	<td align=left>
		<textarea name="evaobser"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_evaobser)%></textarea>
	</td>
</tr>
<tr>
	<td align=right>
		<b>Fecha Pr&oacute;xima Reuni&oacute;n de Seguimiento</b>
	</td>
	<td align=left>
		<input type="text" name="evafecha" size="10" maxlength="10" value="<%=l_evafecha%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.evafecha)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td valign=top align=right colspan=2>
		<a href=# onclick="if (validarfecha(document.datos.evafecha)) {grabar.location='grabar_cierre_COD_eva_00.asp?evldrnro=<%=l_evldrnro%>&evaobser='+escape(Blanquear(document.datos.evaobser.value))+'&evafecha='+document.datos.evafecha.value+'&evareunion='+document.datos.evareunion.value+'&evaetapa=<%=l_evaetapa%>';document.datos.grabado.value='G';}">Grabar</a>
		<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar"-->
<%
cn.Close
set cn = Nothing
%>
</body>
</html>
