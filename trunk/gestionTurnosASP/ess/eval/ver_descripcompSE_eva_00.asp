<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_descripcompSE_eva_00.asp
'Objetivo : Carga de descrips de Competencias 
'Fecha	  : 02-09-2004
'Autor	  : Leticia Amadio
'Modificacion  : 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
on error goto 0

 Dim l_rs, l_rs1
 Dim l_sql
 dim l_cm
 Dim l_filtro
 Dim l_orden

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 Dim l_evaseccnro
 
dim l_evacabnro
dim l_evatevnro
dim l_evaevenro
dim l_evaseccComp
dim l_evldrnroComp
dim l_evatrnro

dim l_compDescrip
dim l_compGrabar
dim l_evatitdesabr

 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")
 l_evaseccnro = Request.QueryString("evaseccnro")
 
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
'set l_rs1 = nothing 

l_evaseccComp =0
l_evldrnroComp = 0
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = " SELECT DISTINCT evadetevldor.evldrnro,evadetevldor.evatevnro, evasecc.evaseccnro "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro= evacab.evacabnro "
l_sql = l_sql & " INNER JOIN evasecc ON evasecc.evaseccnro = evadetevldor.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro "
l_sql = l_sql & " WHERE evacab.evacabnro="& l_evacabnro & " AND evatiposecc.tipsecprog ='carga_calificcompSE_eva_00.asp' "
l_sql = l_sql & "	AND evadetevldor.evatevnro="& cint(cautoevaluador)
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then
	l_evaseccComp = l_rs1("evaseccnro")
	l_evldrnroComp = l_rs1("evldrnro")
	'l_evaevenro = l_rs1("evaevenro")
end if  
l_rs1.Close
set l_rs1 = nothing 

' __________________________________________________________________________________________
' Crear registros de evaresultado para los facnro para el auto evaluador					
' __________________________________________________________________________________________
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT DISTINCT evaseccfactor.evafacnro, evatitulo.evatitnro, evaseccfactor.evaseccnro " ' evaresu.evatrnro, 
l_sql = l_sql & " FROM evaseccfactor " 
l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
l_sql = l_sql & " WHERE evaseccfactor.evaseccnro ="& l_evaseccComp
l_sql = l_sql & " ORDER BY evatitulo.evatitnro "
rsOpen l_rs, cn, l_sql, 0 

set l_cm = Server.CreateObject("ADODB.Command")

if not l_rs.eof then
	l_evatrnro  = "NULL"
	do while not l_rs.eof
  		Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
		l_sql = "SELECT *  FROM  evaresultado "
		l_sql = l_sql & " WHERE evldrnro="&l_evldrnroComp  &" AND evafacnro ="& l_rs("evafacnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then 
			l_sql = "INSERT INTO evaresultado  (evldrnro, evafacnro, evatrnro, evaresudesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnroComp &","& l_rs("evafacnro") & ","& l_evatrnro & ",'')"
			l_cm.activeconnection = Cn 
			l_cm.CommandText = l_sql   
			cmExecute l_cm, l_sql, 0   
		end if
		l_rs1.Close
		set l_rs1=nothing 
	l_rs.MoveNext 
	loop 
end if 
l_rs.Close

' __________________________________________________________________________
' Busca la descripcion del autoevaluador asociada a la competencia 			
' __________________________________________________________________________
sub descripComp (evafacnro, evacabnro,descComp, compGrabar) ',evatevnro
Dim descrip
descrip=""

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaresudesc, evadetevldor.evldrnro "
l_sql = l_sql & " FROM evaresultado "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaresultado.evldrnro "
l_sql = l_sql & " WHERE evaresultado.evafacnro ="& evafacnro & " AND evacabnro="& evacabnro & " AND evatevnro="& cint(cautoevaluador)
l_sql = l_sql & "	AND evaresultado.evldrnro="& l_evldrnroComp
rsOpen l_rs1, cn, l_sql, 0 
if not l_rs1.eof then 
	descrip = l_rs1("evaresudesc") 
end if 
l_rs1.close 
set l_rs1=nothing 

descComp = "<input type=""Text"" name=""resudesc"&evafacnro&l_evldrnroComp &""" value="""& descrip&""" size=""60"" maxlength=""150"" disabled>"
compGrabar="<a href=# onclick=""return false;"">Grabar</a>"
	'compGrabar = compGrabar & "&nbsp;<input class=""rev"" type=""text"" style=""background : #e0e0de;"" readonly disabled name=""grabadoresu"&evafacnro&l_evldrnroComp &""" size=""1"">"		
end sub

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
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

function Validar(fecha)
{
	if (fecha == "") {
		alert("Debe ingresar la fecha .");
		return false;
		}
	else
		{
		return true;
		}
}
 	   

var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}

function Controlar(texto){

    texto.value=Blanquear(texto.value);
 
	if (texto.value.trim()==""){
		alert('Ingrese un Objetivo.');
		texto.focus();
		return false;
	}
	else
		return true;
}	


function ControlarTexto(texto){
	if (texto.value==""){
		alert('Ingrese una Descripción.');
		texto.focus();
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">
<input type="Hidden" name="terminarsecc" value="--">
 <input type="Hidden" name="terminarsecc2" value=""> 

<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr height="20">
	<td colspan="3" align="center">
		<b>DESCRIPCI&Oacute;N DE COMPETENCIAS </b>
	</td>
</tr>
<tr><td colspan="3"><b>AVISO:</b> Se debe completar al menos una descripci&oacute;n para poder terminar la secci&oacute;n.  </td> </tr>
<tr><td colspan="3">&nbsp;</td> </tr>
<tr>
	<th nowrap class="th2">AREAS/COMPETENCIAS </th>
	<th nowrap class="th2">DESCRIPCI&Oacute;N DE LAS COMPETENCIAS </th>
	<th class="th2">&nbsp; </th>
</tr>
<tr><td colspan="3">&nbsp;</td></tr>

<% '
'xxxxxxxxxxxxxxxxxxxxxxxxx
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaseccfactor.evafacnro,evafactor.evafacdesabr, evafactor.evafacdesext, evaseccfactor.orden,evatitulo.evatitnro, evatitulo.evatitdesabr " 
 l_sql = l_sql & " FROM evaseccfactor "
 l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaseccfactor.evafacnro "
 l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
 l_sql = l_sql & " WHERE evaseccfactor.evaseccnro ="& l_evaseccComp
 l_sql = l_sql & " ORDER BY evatitulo.evatitnro, evaseccfactor.orden "  
 rsOpen l_rs, cn, l_sql, 0
'response.write l_sql
 l_evatitdesabr="" 

 do while not l_rs.eof 
	   	' para cada area mostrar su .. evaarea..
	if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then %>
		<tr style="height:20">
			<td align=left valign="middle" colspan="3"><b>AREA: <%=l_rs("evatitdesabr")%> &nbsp;</b> <br> &nbsp;</td>	
		</tr>
<%			l_evatitdesabr = l_rs("evatitdesabr") 
	end if %>
	
	<tr>
		<td valign="top"><%=l_rs("evafacdesabr")%></td>
		<% descripComp l_rs("evafacnro"), l_evacabnro, l_compDescrip, l_compGrabar %>
		<td valign="top"><%= l_compDescrip %></td>
		<td valign=top align=center><%= l_compGrabar%> &nbsp;</td>
	</tr>			
	<% 
	l_rs.Movenext 
loop
l_rs.Close %>

</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<iframe name="terminarsecc" src="termsecc_areasyresultadosSE_eva_00.asp?tipo=carga&evacabnro=<%=l_evacabnro%>&evaseccnro=<%=l_evaseccComp%>&evldrnro=<%=l_evldrnroComp%>&evatevnro=<%=l_evatevnro%>" style="visibility:hidden;width:0;height:0">
</iframe>

</form>	


<%
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
   

</table>

</body>
</html>
