<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 

'Modificado	: 03-12-2004 controlar caracteres raros y priodos dee evaluacion
'Modificado	: 18-03-2005 validar fechas planificacion, seguimiento y evaluacion
'Modificado : 12-10-2005 - Leticia A. - Adaptarlo a Autogestion.


' Variables
' de parametro de entrada
on error goto 0

  Dim l_evaevenro
  Dim l_tipo

' de uso local
  	
' guardan datos del registro de vacpagdesc  
  dim l_evaevedesabr
  dim l_evaevedesext
  dim l_evaevefecha
  dim l_evaevefseg
  dim l_evaevefplan
  dim l_evaevefdesde
  dim l_evaevefhasta
  dim l_evatipnro
  dim l_evaperant
  dim l_evaperact
  dim l_evaperprox
  dim l_evatipevenro
  
' de base de datos
  dim l_rs
  dim l_sql

' parametros entrada
  l_tipo      = Request.QueryString("tipo")
  l_evaevenro = Request.QueryString("evaevenro")	


' BODY ===========================================================================
  select Case l_tipo
	Case "A":
			l_evaevenro		= ""
			l_evaevedesabr  = ""
			l_evaevedesext  = ""
			l_evaevefecha	= date()
			l_evaevefdesde  = date()
			l_evaevefhasta  = date()
			l_evaevefplan   = date()
			l_evaevefseg    = date()
			l_evatipnro		= ""
			l_evaperant		= ""
			l_evaperact		= ""
			l_evaperprox	= ""
			l_evatipevenro	= ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaevenro, evaevedesabr, evaevedesext, evaevefecha,  evaevefplan, evaevefseg, evaevefdesde, evaevefhasta, evatipnro, evaperant, evaperact, evaperprox, evatipevenro FROM  evaevento WHERE evaevento.evaevenro  = " & l_evaevenro 
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_evaevedesabr  = l_rs("evaevedesabr")
			l_evaevedesext  = l_rs("evaevedesext")
			l_evaevefecha   = l_rs("evaevefecha")
			l_evaevefseg   = l_rs("evaevefseg")
			l_evaevefplan   = l_rs("evaevefplan")
			l_evaevefdesde  = l_rs("evaevefdesde")
			l_evaevefhasta  = l_rs("evaevefhasta")
			l_evatipnro		= l_rs("evatipnro")
			l_evaperant		= l_rs("evaperant")
			l_evaperact		= l_rs("evaperact")
			l_evaperprox	= l_rs("evaperprox")
			l_evatipevenro  = l_rs("evatipevenro")
		end if
		l_rs.Close
		set l_rs = nothing
end select

%>
<html>
<head>
<link href="../<%=c_estilo%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Eventos del Ciclo de Gesti&oacute;n del Desempeño<%else%>Evento de Evaluaci&oacute;n<%end if%> - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
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

function Validar_Formulario()
{

	document.datos.evaevedesabr.value = Blanquear(document.datos.evaevedesabr.value);
	document.datos.evaevedesext.value = Blanquear(document.datos.evaevedesext.value);
	
	if (document.datos.evaevedesabr.value.trim() == "")
		alert('Ingrese una descripción.')
	else
	if (document.datos.evatipnro.value == "") 
		alert("Seleccione un Formulario de Evaluación.");
	else
	if (validarfecha(document.datos.evaevefdesde) && validarfecha(document.datos.evaevefhasta) && validarfecha(document.datos.evaevefecha))
	{
		if (document.datos.evaevefdesde.value.trim() == "") 
		alert("Debe ingresar la fecha Desde.");
		else
		if (document.datos.evaevefhasta.value.trim() == "")	
		alert("Debe ingresar la fecha Hasta.");
		else
		if (menorque(document.datos.evaevefhasta.value,document.datos.evaevefdesde.value))
		alert("La Fecha Hasta debe ser posterior a la Fecha Desde.");
		else
		if (document.datos.evaevefecha.value.trim() == "") 
		alert("Debe ingresar la Fecha de Evaluación.");
		else
		if (menor(document.datos.evaevefhasta.value,document.datos.evaevefecha.value))
		alert("La Fecha Hasta debe ser posterior a la Fecha de Evaluación.");
		else
		if (menorque(document.datos.evaevefecha.value,document.datos.evaevefdesde.value))
		alert("La Fecha de Evaluación debe ser posterior a la Fecha Desde.");
		else
		if (document.datos.evatipevenro.value == "") 
		alert("Seleccione un Tipo de Evento");
		else
		<%if ccodelco<>-1 then%>
		if (document.datos.evaperact.value == "") 
		alert("Seleccione un Período a Evaluar.");
		else
		<%end if%>
		<%if ccodelco=-1 then%>
		if (!validarfecha(document.datos.evaevefplan))
		error=true;
		else
		if (!validarfecha(document.datos.evaevefseg))
		error=true;
		else
		if (menorque(document.datos.evaevefseg.value,document.datos.evaevefplan.value))
		alert("La Fecha de Fin de Seguimiento debe ser posterior a la Fecha de Fin de Planificación.");
		else
		if (menorque(document.datos.evaevefecha.value,document.datos.evaevefseg.value))
		alert("La Fecha de Evaluación debe ser posterior a la Fecha de Fin de Seguimiento.");
		else
		if (menorque(document.datos.evaevefplan.value,document.datos.evaevefdesde.value))
		alert("La Fecha de Fin de Planificación debe ser posterior a la Fecha Desde.");
		else
		if (menorque(document.datos.evaevefhasta.value,document.datos.evaevefseg.value))
		alert("La Fecha Hasta debe ser posterior a la Fecha de Fin de Seguimiento.");
		else
		<%end if%>
		{
		var d=document.datos;
		document.valida.location = "evento_evaluacion_06.asp?tipo=<%= l_tipo%>&evaevenro="+document.datos.evaevenro.value + "&evaperact="+document.datos.evaperact.value + "&evaperprox="+document.datos.evaperprox.value+ "&evaevedesabr="+document.datos.evaevedesabr.value;	
		}
	}		
}

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
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

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td align="left" class="th2"><%if ccodelco=-1 then%>Eventos del Ciclo de Gesti&oacute;n del Desempeño<%else%>Evento de Evaluaci&oacute;n<%end if%></td>
	<td align="right" class="th2">&nbsp;</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<form name="datos" action="evento_evaluacion_03.asp?Tipo=<%=l_tipo%>" method="post">
<input type="hidden" name="tipo"	  value="<%=l_tipo%>">
<input type="hidden" name="evaevenro" value="<%=l_evaevenro%>">
<tr>
    <td align="right"><br><b>Descripci&oacute;n:</b>&nbsp;</td>
	<td><br><input name="evaevedesabr"  maxlength=30 size=30 value="<%=l_evaevedesabr%>"></td>
</tr>
          
<tr>
    <td align="right"><br><b>Desc. Extendida:</b>&nbsp;</td>
	<td><br><textarea name="evaevedesext" rows="5" cols="40" maxlength="200"><%= l_evaevedesext %></textarea>&nbsp;</td>
</tr>
<tr>
	<%
	' BUSCAR areas PARA EL <SELECT>
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT evatipnro, evatipdesabr  FROM evatipoeva "
      rsOpen l_rs, cn, l_sql, 0 %>

    <td align="right"><br><b>Formulario de &nbsp;Evaluaci&oacute;n:</b>&nbsp;</td>
	<td>
		<br>
		<select name="evatipnro">
		<option value="">< < Seleccione un Formulario > > </option>
		<%do while not l_rs.eof%>
			<option value=<%=l_rs("evatipnro")%>><%=l_rs("evatipdesabr")%></option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.evatipnro.value='<%=l_evatipnro%>'</script>
	</td>
</tr>
<tr>
	<%
	' BUSCAR areas PARA EL <SELECT>
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT evatipevenro, evatipevedabr FROM evatipoevento "
      rsOpen l_rs, cn, l_sql, 0 %>

    <td align="right"><br><b>Tipo de Evento:</b>&nbsp;</td>
	<td>
		<br>
		<select name="evatipevenro">
		<option value="">< < Seleccione un Tipo de Evento > > </option>
		<%do while not l_rs.eof%>
			<option value=<%=l_rs("evatipevenro")%>><%=l_rs("evatipevedabr")%></option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.evatipevenro.value='<%=l_evatipevenro%>'</script>
	</td>
</tr>
<tr>
    <td colspan="2"><br><b>Fechas del Evento:</b>&nbsp;</td>
</tr>
<tr>
    <td align="right"><b>Desde:</b>&nbsp;</td>
	<td><input type="text" name="evaevefdesde" size="10" maxlength="10" value="<%=l_evaevefdesde%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.evaevefdesde)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
</tr>
<tr>
    <td align="right"><b>Hasta:</b>&nbsp;</td>
	<td><input type="text" name="evaevefhasta" size="10" maxlength="10" value="<%=l_evaevefhasta%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.evaevefhasta)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
</tr>
<%if ccodelco=-1 then%>
<tr>
    <td align="right"><br><b>Fecha Fin Planificaci&oacute;n:</b>&nbsp;</td>
	<td><input type="text" name="evaevefplan" size="10" maxlength="10" value="<%=l_evaevefplan%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.evaevefplan)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
</tr>
<tr>
    <td align="right"><b>Fecha Fin Seguimiento:</b>&nbsp;</td>
	<td><input type="text" name="evaevefseg" size="10" maxlength="10" value="<%=l_evaevefseg%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.evaevefseg)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
</tr>
<%else%>
<input type="hidden" name="evaevefplan">
<input type="hidden" name="evaevefseg">
<%end if%>
<tr>
    <td align="right"><b>Fecha Evaluaci&oacute;n:</b>&nbsp;</td>
	<td><input type="text" name="evaevefecha" size="10" maxlength="10" value="<%=l_evaevefecha%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.evaevefecha)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
</tr>
<%if ccodelco<>-1 then%>
<tr>
    <td colspan="2"><br><b>Per&iacute;odos:</b>&nbsp;</td>
</tr>

<tr>
	<%
	' BUSCAR areas PARA EL <SELECT>
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT evapernro, evaperdesabr, evaperdesde, evaperhasta FROM evaperiodo "
      rsOpen l_rs, cn, l_sql, 0 %>

    <td align="right"><b>Per&iacute;odo a Evaluar:</b>&nbsp;</td>
	<td>
		<select name="evaperact">
		<option value="">< < Seleccione un Período a Evaluar > > </option>
		<%do while not l_rs.eof%>
			<option value=<%=l_rs("evapernro")%>><%=l_rs("evaperdesabr")%>&nbsp;-&nbsp;<%=l_rs("evaperdesde")%>&nbsp;-&nbsp;<%=l_rs("evaperhasta")%></option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.evaperact.value='<%=l_evaperact%>'</script>
	</td>
</tr>
<tr>
	<%
	' BUSCAR areas PARA EL <SELECT>
	  Set l_rs = Server.CreateObject("ADODB.RecordSet")
	  l_sql = "SELECT evapernro, evaperdesabr, evaperdesde, evaperhasta FROM evaperiodo "
      rsOpen l_rs, cn, l_sql, 0 %>

    <td align="right"><b>Pr&oacute;ximo Per&iacute;odo:</b>&nbsp;</td>
	<td>
		<select name="evaperprox">
		<option value="">< < Seleccione un Próximo Período > > </option>
		<%do while not l_rs.eof%>
			<option value=<%=l_rs("evapernro")%>><%=l_rs("evaperdesabr")%>&nbsp;-&nbsp;<%=l_rs("evaperdesde")%>&nbsp;-&nbsp;<%=l_rs("evaperhasta")%></option>
		<%l_rs.MoveNext
		loop
		l_rs.Close
		set l_rs = nothing%>
		</select>
		<script>document.datos.evaperprox.value='<%=l_evaperprox%>'</script>
	</td>
</tr>
<%else%>
<input type="hidden" name="evaperact">
<input type="hidden" name="evaperprox">
<%end if%>
</form>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right">
		<br>
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="0%" height="0%">
</body>
</html>
