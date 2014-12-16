<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: carga_gralobj_eva_00.asp
'Descripción	: Cargar resultado gral de objetivos
'Autor			: 29-12-2004
'Fecha			: CCRossi
'Modificado		: 04-02-2005 L Amadio - modificac para que muestre los resultados de cada objetivo
'										realizado por del autoevaluador y evaluador 
'				: 29-04-2005 - LA. Cambio uso de la constante cevaseccobj --> se usa para tipo de seccion
'           	  13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================
on error goto 0

' Variables
 
' de uso local  
  dim l_evatrnro 
  dim l_evacabnro 
  dim l_ternro  
  dim l_evaluador 
   
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs2, l_rs3
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evacabnro "
  l_sql = l_sql & " FROM evadetevldor "
  l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then
	l_evacabnro=l_rs("evacabnro")
  end if
  l_rs.close
  set l_rs=nothing
  
  	  
' Busco un registro de esta evaluacion que sea de objetivos
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT * "
  l_sql = l_sql & " FROM evadetevldor "
  l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
  l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
  l_sql = l_sql & "        AND evatiposecc.tipsecobj=-1 "
  l_sql = l_sql & " WHERE evadetevldor.evaseccnro = " & l_evaseccnro
  l_sql = l_sql & "   AND evadetevldor.evacabnro = " & l_evacabnro
  rsOpen l_rs, cn, l_sql, 0 
  set l_cm = Server.CreateObject("ADODB.Command") 
  if not l_rs.eof then 
		l_evatrnro  = "null"
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & " FROM  evagralobj "
		l_sql = l_sql & " WHERE evagralobj.evldrnro   = " & l_evldrnro
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			set l_cm = Server.CreateObject("ADODB.Command")  
			l_sql = "INSERT INTO evagralobj "
			l_sql = l_sql & " (evldrnro, evatrnro) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_evatrnro & ")"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs1.Close
		set l_rs1=nothing
  end if
  
'buscar el ternro del EVALUADO --------------------------------------------------------
l_ternro=""
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

'buscar el ternro del EVALUADOR --------------------------------------------------------
l_evaluador =""
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluador "
l_sql = l_sql & " FROM evadetevldor "
l_sql = l_sql & " WHERE evadetevldor.evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_evaluador = l_rs("evaluador")
end if	
l_rs.Close
set l_rs=nothing



' MOSTRAR evaresudes dependiendo del valor que elija como resultado -----
response.write "<script languaje='javascript'>" & vbCrLf
response.write "function Mostrar(evatrnro,evafacnro){ " & vbCrLf
response.write "};" & vbCrLf
response.write "</script>" & vbCrLf
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
.rev
{
	font-size: 10;
	border-style: none;
}
</style>
</head>

<script>

function Controlar(texto,valor){
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
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<form name="datos">
<table border="0" cellpadding="0" cellspacing="1" width="100%">
<%'BUSCAR evagralobj para MODIFICAR resultados ----------------------------
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evagralobj.evatrnro "
   l_sql = l_sql & " FROM evagralobj "
   l_sql = l_sql & " WHERE evldrnro    = " & l_evldrnro
   rsOpen l_rs, cn, l_sql, 0
   
   if  not l_rs.eof then%>
	<tr height="20">
		<td colspan=2 align="right"><b>Evaluaci&oacute;n General de Objetivos:</b>
		&nbsp;
		<%  'BUSCAR la descripcion del resultado  ----------------------------
		    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  evatipresu.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr "
			l_sql = l_sql & " FROM evatipresu  "
			l_sql = l_sql & " WHERE evatrtipo=2 "
			l_sql = l_sql & " order by evatrvalor "
			rsOpen l_rs1, cn, l_sql, 0
			%>
			<select name="evatrnro">	
			<% do while not l_rs1.eof%>
				<option value=<%=l_rs1("evatrnro")%>><%=l_rs1("evatrvalor")%>&nbsp;-&nbsp;<%=l_rs1("evatrdesabr")%></option>
			<%l_rs1.MoveNext
			loop 
			l_rs1.Close
			set l_rs1 = nothing%>
			</select>
			<script>document.datos.evatrnro.value='<%=l_rs("evatrnro")%>'</script>
		</td>
		<td colspan=2 nowrap>
			<a href=# onclick="if (Controlar('texto',document.datos.evatrnro)) {grabar.location='grabar_gralobj_eva_00.asp?evldrnro=<%=l_evldrnro%>&evatrnro='+document.datos.evatrnro.value;document.datos.grabado.value='G'; }">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
		</td>
		<td> &nbsp;</td>
	</tr>
	<tr height="15">
  	        <th align=center colspan=2 class="th2">Descripci&oacute;n</th>
			<th align=center class="th2">&nbsp;</th>
<%			'Buscar Roles
			Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT DISTINCT evatevdesabr "
			l_sql = l_sql & " FROM evaluaobj "
			l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
			l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
			l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
			l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
			l_sql = l_sql & " WHERE evadetevldor.evacabnro= " & l_evacabnro
			l_sql = l_sql & "  AND  evasecc.tipsecnro=" &  cevaseccobj 
			l_sql = l_sql & "  AND (evadetevldor.evatevnro =" & cautoevaluador
			l_sql = l_sql & "   OR  evadetevldor.evatevnro =" & cevaluador & ")"
			l_sql = l_sql & " ORDER BY evatipevalua.evatevdesabr "
			rsOpen l_rs2, cn, l_sql, 0
			do while not l_rs2.eof%>
				<th class="th2"><%=l_rs2("evatevdesabr")%></th>        
			<%l_rs2.MoveNext
			loop
			l_rs2.Close
			set l_rs2=nothing%>
		</tr>
	<%end if
	l_rs.close
	set l_rs = nothing
	
	'BUSCAR OBJETIVOS unicamente
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext,evaobjformed "
	l_sql = l_sql & " FROM evaobjetivo "
	l_sql = l_sql & " INNER JOIN evaluaobj    ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
	l_sql = l_sql & " WHERE evaluaobj.evldrnro=" & l_evldrnro & " AND evadetevldor.evacabnro="& l_evacabnro 
	
	rsOpen l_rs, cn, l_sql, 0
	do while not l_rs.eof %>
	<tr height="10">
        <td align=center valign=middle colspan=2>
			<textarea readonly disabled name="evaobjdext<%=l_rs("evaobjnro")%>"  cols=50 rows=3><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
        <td align=center valign=middle> &nbsp; <!-- saqu e lo de cformed, xq es solo para deloitte -->
			<input readonly disabled name="evaobjformed<%=l_rs("evaobjnro")%>" type=hidden value="<%=trim(l_rs("evaobjformed"))%>">
        </td>
<%			'Buscar RESULTADOS ASOCIADOS A LOS OBJS Y EVALUADORES.  
			Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  Distinct evatipevalua.evatevdesabr, evadetevldor.evatevnro, evaluaobj.evaobjnro "
			l_sql = l_sql & " ,evaluaobj.evatrnro, evatipresu.evatrdesabr  " 'evadetevldor.evldrnro, evadetevldor.evaseccnro
			l_sql = l_sql & " FROM evaluaobj "
			l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
			l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
			l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
			l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
			l_sql = l_sql & " LEFT JOIN evatipresu ON evatipresu.evatrnro = evaluaobj.evatrnro " ' por si alguien no definio resultados.
			l_sql = l_sql & " WHERE evaluaobj.evaobjnro = " & l_rs("evaobjnro")
			l_sql = l_sql & "  AND  evasecc.tipsecnro =" &  cevaseccobj 
			l_sql = l_sql & "  AND (evadetevldor.evatevnro=" & cautoevaluador
			l_sql = l_sql & "       OR evadetevldor.evatevnro=" & cevaluador & ")" 
			l_sql = l_sql & "  AND evacabnro= " & l_evacabnro 
			l_sql = l_sql & " ORDER BY evatipevalua.evatevdesabr"
			'l_sql = l_sql & " GROUP BY evatevdesabr, evadetevldor.evatevnro, evaluaobj.evaobjnro, "
			'l_sql = l_sql & " evaluaobj.evatrnro"
			rsOpen l_rs2, cn, l_sql, 0
			
			do while not l_rs2.eof   
				if l_rs2("evatrdesabr")<> "" then %>
				<td align="center"> 
				<input  type="Text" name="evatrdesabr<%=l_rs2("evatrnro")%>"  readonly disabled value="<%=l_rs2("evatrdesabr")%>" size="35">
				</td>	
		   		<%else %>
				<td align="center">
					<input  type="Text" name="evatrdesabr<%=l_rs2("evatrnro")%>"  readonly disabled value="No se definieron resultados" size="35">
				 </td>
					<script>// document.datos.evatrnro<%'=l_rs("evaobjnro")%><%'=l_rs2("evatevnro")%>.value='<%'=l_rs3("evatrnro")%>'</script>			 
				<% end if 
		  l_rs2.MoveNext
		  loop
		  l_rs2.Close
		  set l_rs2=nothing%>
	</tr>
<%
l_rs.MoveNext
loop
l_rs.Close

cn.Close
Set cn = Nothing
%>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
