<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_cierre_COD_eva_00.asp
'Objetivo : Ver Cierre de una etapa (1: planificacion, 2: seguimiento, 3 evaluacion)
'Fecha	  : 07-02-2005
'Autor	  : CCRossi
'Modificacion: 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

' Variables
' de uso local  
  Dim l_existe  
  Dim l_evareunion
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa
  
  Dim l_texto
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  
' Buscar registros de evafirm para evldrnro y el tipo nota
   l_texto=""
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT *  "
   l_sql = l_sql & " FROM  evacierre "
   l_sql = l_sql & " WHERE evacierre.evldrnro =  " & l_evldrnro
   rsOpen l_rs1, cn, l_sql, 0
'  response.write(l_sql)
   if l_rs1.EOF then
		l_texto= "No hay datos cargados."
   else
  		l_evareunion= l_rs1("evareunion")
  		l_evafecha	= l_rs1("evafecha")
		l_evaobser	= l_rs1("evaobser")
		l_evaetapa	= l_rs1("evaetapa")
   end if
 ' response.write(l_existe)
   l_rs1.Close
   set l_rs1=nothing
   
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cierre - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="<%if trim(l_texto)="" then%><% if l_evareunion=0 then%>document.datos.evareunion.value=0<%else%>document.datos.evareunion.value=-1<%end if%><%end if%>;">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<%if trim(l_texto)<>"" then%>	
<tr>
 	<td colspan=2><%=l_texto%></td>
</tr>
<%else%>
<tr>
 	<td align=right>
 		<b>¿Se realiz&oacute; la reuni&oacute;n de Planificaci&oacute;n?</b>
 	</td>
 	<td align=left>
 		<input type="hidden" name="evareunion" value="<%=l_evareunion%>"><b>
 		<%If l_evareunion = 0 then%> NO <%else%> SI <%End If%></b>
 	</td>
</tr>
<tr>		
	<td align=right>
		<b>Observaci&oacute;n:</b>
	</td>
	<td align=left>
		<textarea style="background : #e0e0de;" readonly name="evaobser"  maxlength=200 size=200 cols=40 rows=5><%=trim(l_evaobser)%></textarea>
	</td>
</tr>
<tr>
	<td align=right>
		<b>Fecha Pr&oacute;xima Reuni&oacute;n de Seguimiento</b>
	</td>
	<td align=left>
		<input style="background : #e0e0de;" readonly type="text" name="evafecha" size="10" maxlength="10" value="<%=l_evafecha%>">
	</td>
</tr>
<tr>
	<td valign=top align=right colspan=2>
		<input class="rev" type="hidden" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
</form>	
<%end if
cn.Close
set cn = Nothing
%>
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
