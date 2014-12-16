<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_areacomRDP_eva_00.asp
'Objetivo : seccin de carga de comentarios por area RDP (mostrar historico RDE)
'Fecha	  : 26-08-2005 
'Autor	  : L.A.	
'Modificado: Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion	
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
on error goto 0
 Dim l_rs, l_rs2,l_rs1
 dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden
 dim i

'locales
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaevenro
 dim l_evaproynro
 dim l_empleado
 dim l_evalrdp

 dim l_evatitnro ' tipo objetivo
 'dim l_cantidad ' cantidad de objetivos de un tipo 
 dim l_evaareacom
 
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_empleg

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 dim l_evaseccnro
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")
 l_evaseccnro = Request.QueryString("evaseccnro")

 if l_orden = "" then
  l_orden = " ORDER BY evatitnro "
 end if


'___________________________________________________
' buscar la evacab   
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro  FROM  evadetevldor WHERE evldrnro ="& l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then 
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
 end if 
 l_rs.close 
 set l_rs=nothing 

 ' fijarse si es RDE o RDP (si evaproynro es Null ent es RDP)
 ' __________________________________________________________
 l_evaproynro = "" 
 l_evalrdp = "NO" 
 Set l_rs = Server.CreateObject("ADODB.RecordSet") 
 l_sql = "SELECT evaevenro, evaproynro, empleado " 
 l_sql = l_sql & " FROM  evadetevldor " 
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro " 
 l_sql = l_sql & " WHERE evldrnro  = " & l_evldrnro 
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.EOF then 
	l_evaproynro = l_rs("evaproynro") 
	l_evaevenro = l_rs("evaevenro") 
	l_empleado = l_rs("empleado") 
 end if 
 
 if l_evaproynro = ""  or  isNull(l_evaproynro) then 
 	l_evalrdp = "SI" 
 end if 
 l_rs.close
 set l_rs=nothing
 
 
 ' Crear los registros de los comentarios .-
 '   XX--XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 '   seleciona todas las areas para el evaluado y el evaluador   
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")  

' selecciono evldrnro para la seccion: Notas Competencias
' selecciono las Areas asociadas a la Evaluacion del empleado (de la seccion: Areas-Competencias)
l_sql = "SELECT DISTINCT evadetevldor.evldrnro, evadetevldor.evaseccnro, evadetevldor.evatevnro, evatitulo.evatitnro "
l_sql = l_sql & " FROM evacab " 
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro "
l_sql = l_sql & " INNER JOIN evadetevldor seccarea ON seccarea.evacabnro = evacab.evacabnro  "  ' join con evacab para sacar ..areas y 
l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evaseccnro = seccarea.evaseccnro "
l_sql = l_sql & " INNER JOIN evafactor ON evaseccfactor.evafacnro = evafactor.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
	' l_sql = l_sql & " LEFT OUTER JOIN evaareacom ON evaareacom.evldrnro = evadetevldor.evldrnro AND evaareacom.evatitnro = evatitulo.evatitnro"
l_sql = l_sql & " WHERE evadetevldor.evaseccnro =" & l_evaseccnro & " AND evadetevldor.evacabnro ="& l_evacabnro 
l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evadetevldor.evldrnro, evadetevldor.evatevnro "
rsOpen l_rs, cn, l_sql, 0 

l_evatitnro = ""

do while not l_rs.eof 
	l_sql = " SELECT * FROM evaareacom WHERE evldrnro = "& l_rs("evldrnro") & " AND evatitnro="& l_rs("evatitnro")
	rsOpen l_rs1, cn, l_sql, 0 
	
	if l_rs1.eof then
		l_sql= "INSERT INTO evaareacom (evatitnro,evldrnro) "
		l_sql = l_sql & " VALUES ("& l_rs("evatitnro")& "," & l_rs("evldrnro") &")"
		
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
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
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

function Controlar(texto,area){

	texto.value=Blanquear(texto.value);
	if (area.value=="")	{
		alert('Seleccione un Area para la Nota .');
		area.focus();
		return false;
	}	
	else
	if (texto.value.trim()=="")	{
		alert('Ingrese una Nota.');
		texto.focus();
		return false;
	} else {	
		return true;	
	}
}	

function ValidarDatos(ponde){
	if (ponde.value=="") {
		alert('Ingrese una Ponderación.');
		ponde.focus();
		return false;
	}	
	else
	if (isNaN(ponde.value)) {
		alert('Ingrese una Ponderación válida.');
		ponde.focus();
		return false;
	}	
	else
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


function ValidarForm (tipo,nota, nroArea,evldrnro, codArea, grabado){
 
if (Controlar(nota,nroArea)){
	nota = escape(nota.value);

	document.datos.evldrnro.value = evldrnro;
	document.datos.evatitnro.value = nroArea; 
	document.datos.evaareacom.value = nota;
	document.datos.evaareacomnro.value = codArea;
	
	document.datos.target ="grabar";
	document.datos.method ="post";
	document.datos.action = "grabar_areacom_eva_00.asp?tipo="+ tipo+'&grabado='+ grabado;
	document.datos.submit();
}
		//'&campo=evaareacom<%'=i%>'

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
        <th align=center colspan=2 class="th2">Notas de Areas de Competencias. </th>
        <th class="th2">&nbsp;</th>
    </tr>
<%
	' seleciona todas las areas para el evaluado y el evaluador 
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT evaareacom, evaareacomnro, evatitulo.evatitnro,evatitdesabr, evatevdesabr, evadetevldor.evatevnro, evadetevldor.evldrnro "
l_sql = l_sql & " FROM evaareacom "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaareacom.evldrnro "
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro "
l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evaareacom.evatitnro "
l_sql = l_sql & " WHERE evacabnro= "& l_evacabnro 
l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evadetevldor.evatevnro "
rsOpen l_rs, cn, l_sql, 0 

' response.write l_sql

l_evatitnro = ""
'l_cantidad  = 0
i=0

do until l_rs.eof
i=i +1
	if l_evatitnro <> l_rs("evatitnro") then 
		l_evatitnro= l_rs("evatitnro") %>
	<tr>
        <td colspan="5"> <br>
			<b><%=l_rs("evatitdesabr")%></b> &nbsp;
			<% if l_evalrdp = "SI" then %>
				<a href=# onclick="Javascript:abrirVentana('detalle_comareasRDE_eva_00.asp?evaevenro=<%=l_evaevenro%>&ternro=<%=l_empleado%>&area=<%=l_rs("evatitnro")%>','',650,400,',scrollbars=yes')" title="Detalle de Comentarios de Areas RDE"> ++</a>	
			<% end if %>	
		</td>
	</tr>
<%	end if%>
    <tr>
        <td align=center valign=top> <%=l_rs("evatevdesabr")%></td>
<%		l_evaareacom = ""
		if l_rs("evaareacom") <> "" then 
			l_evaareacom = unescape(trim(l_rs("evaareacom"))) 
		end if 
%>	
		<% if l_evatevnro <> l_rs("evatevnro") then %>
			<td>
				<textarea name="evaareacom<%=i%>" cols=80 rows=4 readonly><%=l_evaareacom%></textarea>
			</td>
			<td>&nbsp; </td>
		<% else%>
			<td>
				<textarea name="evaareacom<%=i%>" cols=80 rows=4><%=l_evaareacom%></textarea>
			</td>
			<td> 
				<% if not isNull(l_rs("evaareacomnro")) then %>
				<a href=# onclick="ValidarForm('M', document.datos.evaareacom<%=i%>, <%=l_rs("evatitnro")%>, <%=l_rs("evldrnro")%>,<%= l_rs("evaareacomnro")%>, 'document.datos.grabado<%=i%>');">Grabar</a>			
				<% end if %>
				<br> &nbsp;
				<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=i%>" size="1">
			</td>
		<% end if%>
	 </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
%>
</table>



<!-- borrar 
		<a href=# onclick="grabar.location='grabar_areacom_eva_00.asp?tipo=B&evldrnro=<%'=l_evldrnro%>&evaareacomnro=<%'=l_rs("evaareacomnro")%>&evatitnro=<%'=l_rs("evatitnro")%>&evaareacom='+escape(document.datos.evaareacom<%'=i%>.value);document.datos.grabado<%'=i%>.value='B';">Eliminar Nota</a>
		<br>
-->
<%
cn.Close
set cn = Nothing
%>
<input type=hidden name="evldrnro" value="">
<input type=hidden name="evaareacomnro" value="">
<input type=hidden name="evatitnro" value="">
<textarea name="evaareacom"  cols=1 rows=1 style="visibility:hidden;"></textarea>
<input type=hidden name="grabar" value="">
</form>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden; width:0;height:0"> </iframe>
</body>
</html>
