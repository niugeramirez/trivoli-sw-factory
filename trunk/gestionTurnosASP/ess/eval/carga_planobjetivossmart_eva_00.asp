<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'=====================================================================================
'Archivo  : plan_objetivossmart_eva_00.asp
'Objetivo : ABM de planes para objetivos smart 
'Fecha	  : 18-06-2004
'Autor	  : CCRossi
'Fecha	  : 22-10-2004 CCRossi- Cambiar "Modificar" por "Grabar"
'Fecha	  : 28-10-2004 CCRossi- Corregir problema de objetivo creado para otros evaluadores
'Modificacion: 28-10-2004 CCRossi-achicar textareas para que no aparezca scroll en 800x600
'Modificacion: 25-11-04-CCRossi-  control de caraceteres raros
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 Dim l_rs
 Dim l_rs1
 Dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'locales
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaluador 
 dim l_planfecharev
 
'parametros
 Dim l_evldrnro
 
 l_evldrnro = request.querystring("evldrnro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

'Crear los evaplan de cada objetivo--------------------------------------------------

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * FROM evadetevldor "
l_sql = l_sql & " WHERE evldrnro = "& l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evacabnro =l_rs("evacabnro")
	l_evatevnro =l_rs("evatevnro")
	l_evaluador =l_rs("evaluador")
end if	
l_rs.Close
Set l_rs = Nothing

'busco el objetivo asociado al mismo evaluador, mismo evatevnro, misma cabecera.
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadetevldor.evldrnro, evaobjetivo.evaobjnro FROM evadetevldor "
l_sql = l_sql & " INNER JOIN evasecc ON evadetevldor.evaseccnro = evasecc.evaseccnro "
l_sql = l_sql & " INNER JOIN evatiposecc ON evasecc.tipsecnro = evatiposecc.tipsecnro "
l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro=evadetevldor.evldrnro "
l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
l_sql = l_sql & " AND   evatevnro = " & l_evatevnro
l_sql = l_sql & " AND   evaluador = " & l_evaluador
l_sql = l_sql & " AND   tipsecobj=-1" 
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaplan "
	l_sql = l_sql & " WHERE evaobjnro = " & l_rs("evaobjnro")
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	'Response.Write l_sql
	rsOpen l_rs1, cn, l_sql, 0
	if  l_rs1.eof then
		l_rs1.Close
		set l_rs1=nothing
		l_sql= "insert into evaplan (evldrnro,evaobjnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_rs("evaobjnro") &")"
'		Response.Write l_sql
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_rs1.Close
		set l_rs1=nothing
	end if
	
	l_rs.MoveNext
loop	
l_rs.Close
set l_rs=nothing

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
        <th align=center class="th2">Objetivos SMART</th>
        <th align=center class="th2">¿Qu&eacute; necesito alcanzar?</th>
        <th align=center class="th2">¿Cu&aacute;les son los pasos <br>que debo tomar para alcanzar mi objetivo?</th>
        <th align=center class="th2">¿Cu&aacute;ndo cumplir&eacute; mi objetivo?</th>
        <th align=center class="th2">¿Qu&eacute; recursos necesito para cumplir mi objetivo?</th>
        <th align=center class="th2">¿Qui&eacute;n brindar&aacute; apoyo para el logro de mi objetivo?</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext,evaplan.aspectomejorar,evaplan.planaccion,"
l_sql = l_sql & " evaplan.planfecharev, evaplan.recursos,evaplan.ayuda, evaplan.evaplnro "
l_sql = l_sql & " FROM evaplan "
l_sql = l_sql & " INNER JOIN evaobjetivo  ON evaobjetivo.evaobjnro = evaplan.evaobjnro"
l_sql = l_sql & " WHERE evaplan.evldrnro =" & l_evldrnro
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	if trim(l_rs("planfecharev"))="" or isnull(l_rs("planfecharev")) or l_rs("planfecharev")="null" then
		l_planfecharev = date()
	else	
		l_planfecharev = l_rs("planfecharev")
	end if	
	
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
		<td align=center width=15%>
			<b><%=trim(l_rs("evaobjdext"))%></b>
		</td>
        <td align=center width=20%>
			<textarea name="aspectomejorar<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=15 rows=4><%=trim(l_rs("aspectomejorar"))%></textarea>
		</td>
        <td align=center width=20%>
			<textarea name="planaccion<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=15 rows=4><%=trim(l_rs("planaccion"))%></textarea>
		</td>
		<td nowrap width=10%>
			<input type="text" name="planfecharev<%=l_rs("evaobjnro")%>" size="10" maxlength="10" value="<%=l_planfecharev%>">
			<a href="Javascript:Ayuda_Fecha(document.datos.planfecharev<%=l_rs("evaobjnro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
        <td align=center width=20%>
			<textarea name="recursos<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=10 rows=4><%=trim(l_rs("recursos"))%></textarea>
		</td>
        <td align=center width=10%>
			<textarea name="ayuda<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=10 rows=4><%=trim(l_rs("ayuda"))%></textarea>
		</td>
        <td valign=top width=5%>
			<a href=# onclick="if (validarfecha(document.datos.planfecharev<%=l_rs("evaobjnro")%>)) {grabar.location='grabar_planobjetivossmart_eva_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evaobjnro=<%=l_rs("evaobjnro")%>&aspectomejorar='+escape(Blanquear(document.datos.aspectomejorar<%=l_rs("evaobjnro")%>.value))+'&planaccion='+escape(Blanquear(document.datos.planaccion<%=l_rs("evaobjnro")%>.value))+'&planfecharev='+document.datos.planfecharev<%=l_rs("evaobjnro")%>.value+'&recursos='+escape(Blanquear(document.datos.recursos<%=l_rs("evaobjnro")%>.value))+'&ayuda='+escape(Blanquear(document.datos.ayuda<%=l_rs("evaobjnro")%>.value))+'&evaplnro=<%=l_rs("evaplnro")%>';document.datos.grabado<%=l_rs("evaobjnro")%>.value='M';}">Grabar</a>
			<br>
			<a href=# onclick="grabar.location='grabar_planobjetivossmart_eva_00.asp?tipo=B&evaobjnro=<%=l_rs("evaobjnro")%>&evldrnro=<%=l_evldrnro%>&evaplnro=<%=l_rs("evaplnro")%>';document.datos.grabado<%=l_rs("evaobjnro")%>.value='B';">Baja</a>
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

</table>
<input type="Hidden" name="cabnro" value="0">
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->


</form>
</body>
</html>
