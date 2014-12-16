<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%
'================================================================================
'Archivo		: cliente_eva_02.asp
'Descripción	: Abm de Cientes
'Autor			: CCRossi
'Fecha			: 13-12-2004 
'Modificado		: 
'================================================================================


'Datos del formulario
Dim l_evaclinro
Dim l_evaclinom
Dim l_evaclicodext 

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Clientes - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
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
	document.datos.evaclinom.value = Blanquear(document.datos.evaclinom.value);
	document.datos.evaclicodext.value= Blanquear(document.datos.evaclicodext.value);
	
	if (document.datos.evaclinom.value.trim() == "") 
	{
	alert("Debe ingresar un Nombre de Cliente.");
	document.datos.evaclinom.focus();
	}	
	else
	if (document.datos.evaclinom.value.length>60) 
	{
	alert("El Nombre no puede superar 80 caracteres.");
	document.datos.evaclinom.focus();
	}	
	else
	if (document.datos.evaclicodext.value.trim()=="") 
	{
	alert("Debe ingresar un código.");
	document.datos.evaclicodext.focus();
	}	
	else
	if (document.datos.evaclicodext.value.length>60) 
	{
	alert("El Cod.Ext. no puede superar 80 caracteres.");
	document.datos.evaclicodext.focus();
	}	
	else
	{
		var d=document.datos;
		document.valida.location = "cliente_eva_06.asp?tipo=<%= l_tipo%>&evaclinro="+document.datos.evaclinro.value + "&evaclinom="+document.datos.evaclinom.value;	
	}
}

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}


</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		l_evaclinom = ""
	Case "M":
		l_evaclinro = request.querystring("cabnro")
		l_sql = "SELECT evaclinom, evaclicodext "
		l_sql = l_sql & " FROM evacliente "
		l_sql = l_sql  & " WHERE evaclinro = " & l_evaclinro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_evaclinom    = l_rs("evaclinom")
			l_evaclicodext = l_rs("evaclicodext")
		end if
		l_rs.Close
		
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.evaclinom.focus()">
<form name="datos" action="cliente_eva_03.asp?tipo=<%= l_tipo %>" method="post" >

<input type="Hidden" name="evaclinro" value="<%= l_evaclinro %>">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2">Datos del Cliente</td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
    <td align="right"><b>Cod.Ext.:</b></td>
	<td>
		<input type="text" name="evaclicodext" size="21" maxlength="20" value="<%= l_evaclicodext %>">
	</td>
</tr>
<tr>
    <td align="right"><b>Raz&oacute;n Social:</b></td>
	<td>
		<input type="text" name="evaclinom" size="61" maxlength="60" value="<%= l_evaclinom %>">
	</td>
</tr>
<tr height=42>
    <td  colspan="2" valign=top align="right" class="th2">
		<a class="sidebtnABM" href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class="sidebtnABM" href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="0" height="0"></iframe> 
</form>
<%
set l_rs = nothing
%>
</body>
</html>
