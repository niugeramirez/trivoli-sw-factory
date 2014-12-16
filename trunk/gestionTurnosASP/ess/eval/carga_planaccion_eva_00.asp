<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_planaccion_eva_00.asp
'Objetivo : ABM de plan de accion
'Fecha	  : 08-02-2005 * adecuacionpara Codelco
'Autor	  : CCRossi
'Modificacion: 21-03-2005 Poner link Eliminar en rojo y separar
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

Dim l_evldrnro

l_evldrnro = request.querystring("evldrnro")

if l_orden = "" then
  l_orden = " ORDER BY evaplnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
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

function Controlar(texto,rotulo){

    texto.value=Blanquear(texto.value);
 
	if (texto.value.trim()==""){
		alert('Ingrese un ' +rotulo);
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
    <%if ccodelco=-1 then%>
		<th class="th2">Seguimiento</th>
	<%else%>	
        <th class="th2">Aspecto a Mejorar</th>
        <th class="th2">Plan de Accion</th>
     <%end if%> 
        <th class="th2">Fecha <%if ccodelco<>-1 then%>de Revisión<%end if%></th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaplnro, aspectomejorar, planaccion, planfecharev "
l_sql = l_sql & "FROM evaplan "
l_sql = l_sql & "WHERE evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
%>
    <tr>
        <td align=center>
			<textarea name="aspectomejorar<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=40 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
		<%if ccodelco<>-1 then%>
        <td align=center>
			<textarea name="planaccion<%=l_rs("evaplnro")%>"  maxlength=200 size=200 cols=40 rows=4><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		<%else%>
			<input type=hidden name="planaccion<%=l_rs("evaplnro")%>">	
		<%end if%>
        <td valign=middle align=center>
			<input type="text" name="planfecharev<%=l_rs("evaplnro")%>" size="10" maxlength="10" value="<%=l_rs("planfecharev")%>">
			<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev<%=l_rs("evaplnro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td valign=middle align=center>
			<a href=# onclick=" if (validarfecha(document.datos.planfecharev<%=l_rs("evaplnro")%>)) { if (Controlar(document.datos.aspectomejorar<%=l_rs("evaplnro")%>,<%if ccodelco=-1 then%>'seguimiento.'<%else%>'aspecto a mejorar.'<%end if%>)) { grabar.location='grabar_plan_accion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evaplnro=<%=l_rs("evaplnro")%>&aspectomejorar='+escape(document.datos.aspectomejorar<%=l_rs("evaplnro")%>.value)+'&planaccion='+escape(document.datos.planaccion<%=l_rs("evaplnro")%>.value)+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaplnro")%>.value;document.datos.grabado<%=l_rs("evaplnro")%>.value='M';}}">Grabar</a>
			<br>
			<input type="text" class="rev" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evaplnro")%>" size="1">
			<br>
			<%if ccodelco=-1 then%>
			<a href=# style="color:red;" onclick="if (confirm('¿ Desea Eliminar el Seguimiento?')==true) { grabar.location='grabar_plan_accion_00.asp?tipo=B&evaplnro=<%=l_rs("evaplnro")%>&evldrnro=<%=l_evldrnro%>&aspectomejorar='+document.datos.aspectomejorar<%=l_rs("evaplnro")%>.value+'&planaccion='+document.datos.planaccion<%=l_rs("evaplnro")%>.value+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaplnro")%>.value};document.datos.grabado<%=l_rs("evaplnro")%>.value='B';">Eliminar</a>
			<%else%>			
			<a href=# style="color:red;" onclick="grabar.location='grabar_plan_accion_00.asp?tipo=B&evaplnro=<%=l_rs("evaplnro")%>&evldrnro=<%=l_evldrnro%>&aspectomejorar='+document.datos.aspectomejorar<%=l_rs("evaplnro")%>.value+'&planaccion='+document.datos.planaccion<%=l_rs("evaplnro")%>.value+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaplnro")%>.value;document.datos.grabado<%=l_rs("evaplnro")%>.value='B';">Eliminar</a>
			<%end if%>
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
<tr>
    <td align=center>
		<textarea name="aspectomejorar"  maxlength=200 size=200 cols=40 rows=4></textarea>
	</td>
	<%if ccodelco<>-1 then%>
    <td align=center>
		<textarea name="planaccion"  maxlength=200 size=200 cols=40 rows=4></textarea>
	</td>
	<%else%>
		<input type=hidden name="planaccion">	
	<%end if%>
        
	<td valign=middle align=center>
		<input type="text" name="planfecharev" size="10" maxlength="10" value="<%=Date()%>">
		<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
	<td valign=middle align=center>
		<a href=# onclick="if (validarfecha(document.datos.planfecharev)) {if (Controlar(document.datos.aspectomejorar,<%if ccodelco=-1 then%>'seguimiento.'<%else%>'aspecto a mejorar.'<%end if%>)) {grabar.location='grabar_plan_accion_00.asp?tipo=A&evldrnro=<%=l_evldrnro%>&aspectomejorar='+escape(document.datos.aspectomejorar.value)+'&planaccion='+escape(document.datos.planaccion.value)+'&planfecharev='+document.datos.planfecharev.value;document.datos.grabado.value='G';}}">Grabar</a>
		<br>
		<input type="text" class="rev" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
	</td>
</tr>

</table>
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

<input type="Hidden" name="cabnro" value="0">
</form>
</body>
</html>
