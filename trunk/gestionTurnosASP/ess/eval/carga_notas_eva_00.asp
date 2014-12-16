<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Modificado: 24-11-2004 CCRossi - Control de caracteres raros
'Modificado: 17-02-2005 CCRossi - Cambiar campo dext por desabr
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.

' Variables
' de parametros entrada
  
' de uso local  
    
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm
  dim l_armado
  
' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

  
' Crear registros de evaNOTAS para evldrnro y el tipo nota
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evatnnro FROM evaseccnota WHERE evaseccnota.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0

  set l_cm = Server.CreateObject("ADODB.Command")  
  do while not l_rs.eof
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT * FROM  evanotas "
		l_sql = l_sql & " WHERE evanotas.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " AND   evanotas.evatnnro  = " & l_rs("evatnnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evanotas (evldrnro, evatnnro, evanotadesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," &  l_rs("evatnnro")
			l_sql = l_sql &  ",'')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs.MoveNext
		l_rs1.Close
  loop
  l_rs.Close
  set l_rs=nothing 
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Notas de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function Blanquear(texto){
 var  aux;
 aux = replaceSubstring(texto,"'","");
 aux = replaceSubstring(aux,"´","");
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

function validar(texto,ubic)
{
alert("entro")
}

</script>
<style>
.rev
{
	font-size: 10;
	border-style: none;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">
<input type="Hidden" name="terminarsecc" value="SI">
<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="5" align="left" class="th2"><%if ccodelco=-1 then%>Conclusiones<%else%>Carga de Notas de Evaluaci&oacute;n<%end if%></td>
<tr>
<tr style="border-color :CadetBlue;">
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<%if ccodelco<>-1 then%>
<tr style="border-color :CadetBlue;">
	<td><b>Tipo de Nota</b></td>
	<td><b>Nota</b></td>
	<td>&nbsp;</td>
</tr>
<%end if%>	
<%'BUSCAR evaNotas para MODIFICAr resultados ------------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evldrnro, evanotas.evatnnro, evanotadesc, evatndesabr, evatndesext,evaseccnota.orden "
   l_sql = l_sql & "FROM evanotas "
   l_sql = l_sql & "INNER JOIN evaseccnota ON evaseccnota.evatnnro = evanotas.evatnnro "
   l_sql = l_sql & "INNER JOIN evatiponota ON evatiponota.evatnnro = evanotas.evatnnro "
   l_sql = l_sql & "WHERE evaseccnota.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evanotas.evldrnro      = " & l_evldrnro
   l_sql = l_sql & " ORDER BY evaseccnota.orden "
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof %>
   <tr>
		<td valign=top><%=l_rs("evatndesabr")%></td>
		<td>
		<textarea name="evanotadesc<%=l_rs("evatnnro")%>"  maxlength=255 size=255 cols=40 rows=6><%=trim(l_rs("evanotadesc"))%></textarea>
		</td>
			<td valign=top><a href=# onclick="grabar.location='grabar_notas_evaluacion_00.asp?evatnnro=<%=l_rs("evatnnro")%>&evldrnro=<%=l_evldrnro%>&evanotadesc='+ escape(Blanquear(document.datos.evanotadesc<%=l_rs("evatnnro")%>.value));document.datos.grabado<%=l_rs("evatnnro")%>.value='G';">Grabar</a>
			<input type="text" class="rev" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evatnnro")%>" size="1">
		</td>
    </tr>
  <%l_rs.Movenext
  loop
  l_rs.Close
  set l_rs=nothing

cn.Close
set cn = Nothing
%>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
