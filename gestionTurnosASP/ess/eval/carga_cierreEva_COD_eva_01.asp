<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script>
	var textonotafinal="";
	textonotafinal = textonotafinal + "1.0 -- 1.9: No cumple \n\n"
	textonotafinal = textonotafinal + "2.0 -- 2.9: Cumple en parte\n\n"
	textonotafinal = textonotafinal + "3.0 -- 3.9: Cumple\n\n"
	textonotafinal = textonotafinal + "4.0 -- 4.9: Supera lo comprometido\n\n"
	textonotafinal = textonotafinal + "       5.0: Excelencia"
	
</script>
<% 
'=====================================================================================
'Archivo  : carga_cierreEva_COD_eva_01.asp
'Objetivo : Cierre de una etapa (1: planificacion, 2: seguimiento, 3 evaluacion)
'Fecha	  : 15-02-2005
'Autor	  : CCRossi
'Modificacion: 18-03-2005 CCRossi-Agregar ayuda de nota y validacion de nota
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
' Datos del garante
'=====================================================================================

' Variables
' de uso local  
  Dim l_existe  
  Dim l_evareunion
  Dim l_evafecha
  Dim l_evaobser
  Dim l_evaetapa

  dim l_notafinal
  
  Dim l_evacabnro
  Dim l_lista
  Dim l_primero
  
  dim l_caracteristica  
  dim l_nombre
       
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_rs2
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_evaseccnro
  Dim l_vacio
     
' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_vacio = Request.QueryString("mostrar")
  l_evaetapa = 3

if l_vacio="1" then
	'buscar la evacab
	 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	 l_sql = "SELECT evacabnro  "
	 l_sql = l_sql & " FROM  evadetevldor "
	 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
	 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
	 rsOpen l_rs1, cn, l_sql, 0
	 if not l_rs1.EOF then
		l_evacabnro = l_rs1("evacabnro")
	 end if
	 l_rs1.close
	 set l_rs1=nothing

	'Crear registros de evacierre 
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
			l_sql = l_sql & "(evldrnro,evaetapa) "
			l_sql = l_sql & "values (" & l_rs("evldrnro") &","&l_evaetapa &")"
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
	<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
	<script>
	function Validar(alcanza){
	
	if (alcanza.disabled==false)
	{
		if (alcanza.value=="") 
		{
			alert('Ingrese una Nota.');
			alcanza.focus();
			return false;
		}	
		else
		if (isNaN(alcanza.value))
		{
			alert('Ingrese una Nota válida.\nUtilice punto (".") para decimales.');
			alcanza.focus();
			return false;
		}	
		else	
		if (!validanumero(alcanza, 1, 1)) 
		{
			alert('Ingrese una Nota con 1 entero y 1 decimal, con "." como separador de decimales.');
			alcanza.focus();
			return false;
		}
		else
		if ((Number(alcanza.value) >5) || (Number(alcanza.value)<1)) 
		{
			alert('Ingrese una Nota entre 1 y 5.');
			alcanza.focus();
			return false;
		}	
		else
			return true;		
} // no es readonly
	} //function 
	
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

	<style>
	.rev
	{
		font-size: 11;
		border-style: none;
	}
	</style>

	</head>
	<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >

	<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr>		
		<td align="center" colspan="4"><b>Mediador Garante SOLO EN CASO DE DESACUERDO:</b></td>
	</tr>
		
	<form name="datos">
	<%
	' SI EL GARANTE ESTA HABILITADO (POR DESACUERDO) MUESTRO ESTA INFO 
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evacierre.evldrnro, evareunion, evaacuerdo, evadetevldor.evatevnro, evatevdesabr , evaobser, "
	l_sql = l_sql & " empleado.empleg"
	l_sql = l_sql & " FROM  evacierre "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evacierre.evldrnro"
	l_sql = l_sql & "		 AND ( evadetevldor.evatevnro<> " & cautoevaluador
	l_sql = l_sql & "		   AND evadetevldor.evatevnro<> " & cevaluador & ")"
	l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
	l_sql = l_sql & " WHERE evacierre.evldrnro IN (" & l_lista & ")"
	rsOpen l_rs1, cn, l_sql, 0
	l_primero=-1
	do while not l_rs1.eof
		'response.write l_evldrnro & "<BR>"	
		'response.write l_rs1("evldrnro") & "<BR>"
		if cdbl(l_evldrnro) <> cdbl(l_rs1("evldrnro")) then
			l_caracteristica = "readonly disabled style='background : #e0e0de;'"
			l_nombre = l_evldrnro
		else	
			l_caracteristica = ""
			l_nombre = ""
		end if
		%>
		<tr>		
			<td align=right >
				<b>Comentarios <%=l_rs1("evatevdesabr")%>:</b>
			</td>
			<td align=left>
				<textarea <%=l_caracteristica%> name="evaobser<%=l_evldrnro%>" maxlength=200 size=200 cols=70 rows=3><%=trim(l_rs1("evaobser"))%></textarea>
			</td>
		<%
		l_rs1.Movenext
	Loop
	l_rs1.close
	set l_rs1=nothing
	
	'BUSCAR Nota Final
	l_notafinal=""
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT puntajemanual "
	l_sql = l_sql & " FROM  evacab"
	l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro
	rsOpen l_rs2, cn, l_sql, 0
	if not l_rs2.EOF then
		l_notafinal= l_rs2("puntajemanual")
	end if
	l_rs2.close
	set l_rs2=nothing
	
	if trim(l_notafinal)="" or isnull(l_notafinal) then
		'Calcular Nota Final
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
		do while not l_rs2.EOF 
			l_notafinal= l_notafinal + cdbl(l_rs2("puntaje")) * cdbl(l_rs2("evatipopor")) / 100
			l_rs2.Movenext
		loop
		l_rs2.close
		set l_rs2=nothing
	end if%>
	
	
		<td align=right ><b>NOTA FINAL:</b></td>
		<td align=left><input <%=l_caracteristica%> onblur="Validar(this);" type="text" name="notafinal" maxlength=3 size=3 value="<%=l_notafinal%>">
		<a href=# onclick="alert(textonotafinal);">?</a></td>
	</tr>

	</form>	
	</table>

	<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
	</iframe>

	</body>
	</html>
<%else%>

<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Cierre de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<form name="datos">
	<tr>		
		<td colspan=2><br><br><br><br><br><br><br></td>
	</tr>
</form>	
</table>
</body>
</html>
<%end if

cn.Close
set cn = Nothing

%>