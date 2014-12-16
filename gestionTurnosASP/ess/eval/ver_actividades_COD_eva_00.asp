<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_atividades_COD_eva_00.asp
'Objetivo : Ver de actividades or objetivos (mas de una)
'Fecha	  : 07-02-2005
'Autor	  : CCRossi
'Modificacion: 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

Dim l_evaluador
Dim l_evacabnro
Dim l_evldrnro

Dim l_evatipobjnro

l_evldrnro = request.querystring("evldrnro")

if l_orden = "" then
  l_orden = " ORDER BY evaplnro "
end if
'Busco evaluador...
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaluador, evacabnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaluador = l_rs("evaluador")
	l_evacabnro = l_rs("evacabnro")
 end if
 l_rs.close
 set l_rs=nothing

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
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

function Controlar(texto){

    texto.value=Blanquear(texto.value);
 
	if (texto.value.trim()==""){
		alert('Ingrese un Aspecto a Mejorar.');
		texto.focus();
		return false;
	}
	else
		return true;
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr style="border-color :CadetBlue;">
        <th class="th2">Compromiso</th>
        <th class="th2">Actividades</th>
        <th class="th2">Fecha Comprometida</th>
        <th class="th2">Aproyo Comprometido</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evatipobjnro, evatipobjdabr, evaplnro, aspectomejorar, planaccion, planfecharev, evaobjdext, evaplan.evaobjnro  "
l_sql = l_sql & " FROM evaplan "
l_sql = l_sql & " INNER JOIN evaobjetivo ON evaplan.evaobjnro=evaobjetivo.evaobjnro "
l_sql = l_sql & " INNER JOIN evatipoobj  ON evatipoobj.evatipobjnro=evaobjetivo.evatipobjnro "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro=evaobjetivo.evaobjnro "
l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " INNER JOIN evadetevldor ON evaluaobj.evldrnro=evadetevldor.evldrnro "
l_sql = l_sql & "		AND evadetevldor.evaluador= " & l_evaluador
l_sql = l_sql & "		AND evadetevldor.evacabnro= " & l_evacabnro
l_sql = l_sql & " WHERE evaplan.evldrnro =" & l_evldrnro
l_sql = l_sql & " ORDER BY evaobjetivo.evatipobjnro, evatipobjdabr, evaplnro, aspectomejorar, planaccion, planfecharev, evaobjdext, evaplan.evaobjnro  "
rsOpen l_rs, cn, l_sql, 0
'Response.Write l_sql
l_evatipobjnro="" 
if l_rs.eof then%>
	<tr>
		<td colspan="5">No hay Actividades cargadas.</td>
	</tr>
<%
end if
do until l_rs.eof
%>
    <tr>
		<%if l_evatipobjnro <> l_rs("evatipobjnro") then
			l_evatipobjnro= l_rs("evatipobjnro") 
			%>
			<tr>
				<td colspan="5"><b><%=l_rs("evatipobjdabr")%></b></td>
			</tr>
		<%end if%>
		<td align=left><%=l_rs("evaobjdext")%></td>
        <td>
			<textarea style="background : #e0e0de;" readonly name="aspectomejorar<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
        <td>
			<input style="background : #e0e0de;" readonly type="text" name="planfecharev<%=l_rs("evaplnro")%>" size="10" maxlength="10" value="<%=l_rs("planfecharev")%>">
		</td>
        <td>
			<textarea style="background : #e0e0de;" readonly name="planaccion<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		<td valign=top>
			<input type="hidden" class="rev" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evaplnro")%>" size="1">
		</td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
%>
</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
<%cn.Close
set cn = Nothing%>
