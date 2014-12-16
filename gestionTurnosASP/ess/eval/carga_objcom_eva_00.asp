<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_objcom_eva_00.asp
'Objetivo : seccion de carga de comentarios por objetivos
'Fecha	  : 10-01-2005
'Autor	  : Leticia A.
'Modificacion: 15-04-2005 - un comentario por obj (permitir ingresar mas texto por comentario)
'			: 29-04-2005 - LA. Cambio uso de la constante cevaseccobj --> se usa para tipo de seccion
' 			: 19-07-2005 - L.A. - - crear todos los registros para los cometarios.
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

on error goto 0
 Dim l_rs, l_rs2, l_rs1
 Dim l_sql
 Dim l_cm
 Dim l_filtro
 Dim l_orden
 
 dim i, j
 
'locales
 dim l_evacabnro 
 dim l_evatevnro 

Dim l_evatitnro ' tipo objetivo
Dim l_evaobjnro
dim l_cantidad ' cantidad de objetivos de un tipo  
Dim l_objetivo
Dim l_evaobjcom

 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 					' lo saque!!!!!
 dim l_empleg

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if


'___________________________________________________________________________________
'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
 end if
 l_rs.close
 set l_rs=nothing
 
 
 ' _________________________________________________________________________________
 ' Crea los registros de comentarios de Objetivo (evaobjcom) 
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")  

l_sql = "SELECT distinct evaobjetivo.evaobjnro, evadetevldor.evatevnro,evadetevldor.evldrnro"
l_sql = l_sql & " FROM evaobjetivo "
	' evaluaobj --> dado que se crea un reg ahi, por cada objetivo que se crea!!
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
	' l_sql = l_sql & " LEFT  JOIN evaobjcom ON evaobjcom.evldrnro=evadetevldor.evldrnro AND evaobjcom.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " WHERE  evasecc.tipsecnro <>" &  cevaseccobj 
l_sql = l_sql & "   AND (evadetevldor.evatevnro=" & cautoevaluador
l_sql = l_sql & "          OR evadetevldor.evatevnro=" & cevaluador & ")"
l_sql = l_sql & "   AND evacabnro=" & l_evacabnro 
l_sql = l_sql & " ORDER BY evaobjetivo.evaobjnro,evadetevldor.evatevnro"
rsOpen l_rs, cn, l_sql, 0 
' response.write l_sql & "<br>"

do while not l_rs.eof 
	l_sql = " SELECT * FROM evaobjcom WHERE evldrnro = "& l_rs("evldrnro") & " AND evaobjnro="& l_rs("evaobjnro")
	rsOpen l_rs1, cn, l_sql, 0 
	
	if l_rs1.eof then 
		l_sql= "INSERT INTO evaobjcom (evaobjnro, evldrnro) "
		l_sql = l_sql & " VALUES ("&l_rs("evaobjnro") & "," & l_rs("evldrnro") &")"
		
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if 
	l_rs1.Close 
	
l_rs.MoveNext
loop

l_rs.Close
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>

<script>
String.prototype.trim = function() {

 // skip leading and trailing whitespace
 // and return everything in between
  var x=this;
  x=x.replace(/^\s*(.*)/, "$1");
  x=x.replace(/(.*?)\s*$/, "$1");
  return x;
}

function Controlar(texto,nro){

	texto.value=Blanquear(texto.value);

	if (nro.value==""){
		alert('Seleccione un Objetivo para la Nota .');
		area.focus();
		return false;
	}	
	else
	if (texto.value.trim()=="")	{
		alert('Ingrese una Nota.');
		texto.focus();
		return false;
	}else{	
		return true;	
	}
}	

function ValidarDatos(ponde){
	if (ponde.value=="") 	{
		alert('Ingrese una Ponderación.');
		ponde.focus();
		return false;
	}	
	else
	if (isNaN(ponde.value)) {
		alert('Ingrese una Ponderación válida.');
		ponde.focus();
		return false;
	}else
		return true;
		
}
 

var jsSelRow = null;

function Deseleccionar(fila){
 fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
 if (jsSelRow != null) {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
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


function ValidarForm (tipo,nota, nroObj,evldrnro, codObj, grabado){
 
if (Controlar(nota,nroObj)){
	nota = escape(nota.value);

	document.datos.evldrnro.value = evldrnro;
	document.datos.evaobjnro.value = nroObj;
	document.datos.evaobjcom.value = nota;
	document.datos.evaobjcomnro.value = codObj;
	
	document.datos.target ="grabar";
	document.datos.method ="post";
	document.datos.action = "grabar_objcom_eva_00.asp?tipo="+ tipo+'&grabado='+ grabado;
	document.datos.submit();
}
		
}
			 

</script>
<style>
.blanc {
	font-size: 10;
	border-style: none;
	background : transparent;
}
.rev {
	font-size: 10;
	border-style: none;
}
</style>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" method="post">
<input type="Hidden" name="terminarsecc" value="SI">

<table>
    <tr>
        <th align=center colspan=2 class="th2">Notas de Objetivos </th>
        <th class="th2">&nbsp;</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjdext, evatipevalua.evatevdesabr, evadetevldor.evatevnro,evadetevldor.evldrnro, evaobjcom.evaobjcom, evaobjcomnro"
l_sql = l_sql & " FROM evaobjcom " 
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaobjcom.evldrnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaobjcom.evaobjnro "
l_sql = l_sql & " WHERE evacabnro=" & l_evacabnro 
l_sql = l_sql & " ORDER BY evaobjetivo.evaobjnro, evatipevalua.evatevdesabr"
rsOpen l_rs, cn, l_sql, 0 
 
l_objetivo = "" 
i = 0 
j = 0 

if not l_rs.eof then
	
	do until l_rs.eof 
		i= i + 1
		l_evaobjcom=""
		if not isNull(l_rs("evaobjcom")) and l_rs("evaobjcom")<> "" then
			l_evaobjcom = unescape(trim(l_rs("evaobjcom")))
		end if
		if l_objetivo <> l_rs("evaobjnro") then 
			j=j+1 %>
		<tr>
			<td valign="center"><b>Objetivo <%=j%> </b></td>
			<td colspan="2"><%=l_rs("evaobjdext")%></td>
		</tr>
		<% end if %>
		<tr>
			<td valign="top"><%=l_rs("evatevdesabr")%></td>
			<% if l_evatevnro <> l_rs("evatevnro") then %>
				<td>
				<textarea name="evaobjcom1<%=i%>"  cols=80 rows=4 readonly><%=l_evaobjcom%></textarea>
				</td> 
				<td>&nbsp; </td>
			<% else%>
				<td>
				<textarea name="evaobjcom<%=i%>" cols=80 rows=4><%=l_evaobjcom%></textarea> 
				</td>
				<td> 
				<a href=# onclick="ValidarForm('M', document.datos.evaobjcom<%=i%> ,<%=l_rs("evaobjnro")%>, <%=l_rs("evldrnro")%>,<%= l_rs("evaobjcomnro")%>, 'document.datos.grabado<%=i%>');">Grabar</a>
				<br> &nbsp;
				<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=i%>" size="1">
				</td>
			<% end if %>
		</tr>
<% 
		l_objetivo = l_rs("evaobjnro")
	l_rs.MoveNext
	
	loop
	
else %>
	<tr><td colspan="3"> No se han definido Objetivos. </td></tr>
<%end if

l_rs.Close
set l_rs = nothing%>

</table>

<%
cn.Close
set cn = Nothing
%>

<input type=hidden name="evldrnro" value="">
<input type=hidden name="evaobjcomnro" value="">
<input type=hidden name="evaobjnro" value="">
<textarea name="evaobjcom"  cols=1 rows=1 style="visibility:hidden;"></textarea>
<input type=hidden name="grabar" value="">
</form>

 <!-- <iframe name="grabar" style="width:200;height:200">  -->
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0"> 
</body>
</html>
