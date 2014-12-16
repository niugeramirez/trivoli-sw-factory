<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_objetivos_eva_00.asp
'Objetivo : ABM de objetivos de evaluacion
'Fecha	  : 06-05-2004
'Autor	  : CCRossi
'Modificacion  : 13-07-2004 - CCRossi - Controlar que cargue Objetivo cuando guarda o modifica
'Modificacion: 19-11-04-CCRossi- titulo Gestion y control de caraceteres raros
'Modificacion: 24-11-04-CCRossi- control de caraceteres raros
'Modificacion: 29-12-04-CCRossi- sacar campos para Deloitte 
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

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
<table>
    <tr>
        <th align=center class="th2">Descripci&oacute;n</th>
        <%if cformed=-1 then%>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <%else%>
        <th align=center class="th2">&nbsp;</th>
        <%end if%>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<input type="Hidden" name="terminarsecc" value="SI">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evldrnro "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
        <td align=center>
			<textarea name="evaobjdext<%=l_rs("evaobjnro")%>"  size=200 cols=50 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
        <td align=center>
			<%if cformed=-1 then%>
			<textarea name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjformed"))%></textarea>
			<%else%>
			<input name="evaobjformed<%=l_rs("evaobjnro")%>" type=hidden value="<%=trim(l_rs("evaobjformed"))%>">
			<%end if%>
		</td>
        <td valign=top>
			<a href=# onclick="if (Controlar(document.datos.evaobjdext<%=l_rs("evaobjnro")%>)) {  grabar.location='grabar_objetivos_eva_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evapernro=<%=l_evapernro%>&evaobjnro=<%=l_rs("evaobjnro")%>&evaobjdext='+escape(document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value)+'&evaobjformed='+escape(Blanquear(document.datos.evaobjformed<%=l_rs("evaobjnro")%>.value));document.datos.grabado<%=l_rs("evaobjnro")%>.value='M'; }">Grabar</a>
			<br>
			<a href=# onclick="grabar.location='grabar_objetivos_eva_00.asp?tipo=B&evaobjnro=<%=l_rs("evaobjnro")%>&evapernro=<%=l_evapernro%>&evldrnro=<%=l_evldrnro%>&evaobjdext='+escape(Blanquear(document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value))+'&evaobjformed='+escape(Blanquear(document.datos.evaobjformed<%=l_rs("evaobjnro")%>.value));document.datos.grabado<%=l_rs("evaobjnro")%>.value='B';">Eliminar Obj.</a>
			<br>
			<input type="text" readonly disabled name="grabado<%=l_rs("evaobjnro")%>" size="1">
		</td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
    <tr onclick="Javascript:Seleccionar(this,0)">
        <td align=center>
			<textarea name="evaobjdext" size=200 cols=50 rows=4></textarea>
		</td>
        <td align=center>
        <%if cformed=-1 then%>
			<textarea name="evaobjformed"  maxlength=200 size=200 cols=30 rows=4></textarea>
		<%else%>
			<input name="evaobjformed" type=hidden>
		<%end if%>
		</td>
		<td valign=top>
			<a href=# onclick="if (Controlar(document.datos.evaobjdext)) {  grabar.location='grabar_objetivos_eva_00.asp?tipo=A&evapernro=<%=l_evapernro%>&evldrnro=<%=l_evldrnro%>&evaobjdext='+escape(document.datos.evaobjdext.value)+'&evaobjformed='+escape(document.datos.evaobjformed.value);document.datos.grabado.value='G'; }">Grabar</a>
			<br>
			<input type="text" readonly disabled name="grabado" size="1">
		</td>
    </tr>

</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
