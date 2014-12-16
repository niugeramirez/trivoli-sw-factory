<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: carga_estructuras_eva_00.asp
'Descripción	: Cargar estructuras segun tipo de evaseccestr
'Autor			: 16-09-2004
'Fecha			: CCRossi
'Modificado		: 30-09-2004 CCRossi arreglar el select porque mostraba cosas mezcladas
'Modificar		: 25-11-2004-CCRossi- Controlar caracteres raros
'            	  13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================
 Dim l_rs
 Dim l_rs1
 Dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden

'locales
 dim l_puntacion
 dim l_puntajemanual
 dim l_puntaje
 
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaseccnro 
 dim l_estrnro
  
'parametros
 Dim l_evldrnro
 
 l_evldrnro = request.querystring("evldrnro")
 
 if l_orden = "" then
  l_orden = " ORDER BY orden "
 end if

'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro, evatevnro, evaseccnro  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
	l_evaseccnro = l_rs("evaseccnro")
 end if
 l_rs.close
 set l_rs=nothing

'Crear registros de evaNOTAS para evldrnro y el tipo nota
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT tenro "
  l_sql = l_sql & "FROM evaseccestr "
  l_sql = l_sql & "WHERE evaseccestr.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0
  set l_cm = Server.CreateObject("ADODB.Command")  
  do while not l_rs.eof
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & "FROM  evaestructuras "
		l_sql = l_sql & "WHERE evaestructuras.tenro  = " & l_rs("tenro")
		l_sql = l_sql & "AND   evaestructuras.evacabnro  = " & l_evacabnro
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evaestructuras "
			l_sql = l_sql & "(evacabnro, tenro ) "
			l_sql = l_sql & " VALUES (" & l_evacabnro &","& l_rs("tenro") &")"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		else
			l_estrnro=l_rs1("estrnro")
		end if
		l_rs1.close
		set l_rs1=nothing
		
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & "FROM  evaluaestr "
		l_sql = l_sql & "WHERE evaluaestr.tenro  = " & l_rs("tenro")
		l_sql = l_sql & "AND   evaluaestr.evldrnro  = " & l_evldrnro
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			if trim(l_estrnro)<>"" then
				l_sql = "INSERT INTO evaluaestr "
				l_sql = l_sql & "(evldrnro, tenro,estrnro) "
				l_sql = l_sql & " VALUES (" & l_evldrnro &","& l_rs("tenro") &","& l_estrnro & ")"
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
			else
				l_sql = "INSERT INTO evaluaestr "
				l_sql = l_sql & "(evldrnro, tenro) "
				l_sql = l_sql & " VALUES (" & l_evldrnro &","& l_rs("tenro") & ")"
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
			end if	
		end if
		l_rs1.Close
		set l_rs1=nothing
		
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

		
function ValidarDatos(estructura,desde,hasta,dext)
{
	dext.value = Blanquear(dext.value);
	
	if (estructura.registro.grupo.value == "-1") 
	{
		alert("Debe seleccionar una Estructura.");
		estructura.registro.grupo.focus();
	}	
	else
	if (desde.value.trim() == "") 
	{
		alert("Ingrese una Fecha Desde.");
		desde.focus();
	}	
	else
	if (!validarfecha(desde)) 
	{
		alert('Ingrese una Fecha Desde válida.');
		desde.focus();
		return false;
	}	
	else
	if (hasta.value.trim() == "") 
	{
		alert("Ingrese una Fecha Hasta.");
		hasta.focus();
	}	
	else
	if (!validarfecha(hasta)) 
	{
		alert('Ingrese una Fecha Hasta válida.');
		desde.focus();
		return false;
	}	
	else
	if (!menorque(desde.value,hasta.value)) 
	{
		alert('La Fecha Desde debe ser menor o igual que la Fecha Hasta.');
		desde.focus();
		return false;
	}	
	else
	if (dext.value.length >255) 
	{
		alert('La Observación no puede superar 255 caracteres.');
		dext.focus();
		return false;
	}	
	else
		return true;
		
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
    <tr>
        <th align=center class="th2">Tipos de Estructura</th>
        <th align=center class="th2">Estructura</th>
        <th align=center class="th2">Fecha Desde</th>
        <th align=center class="th2">Fecha Hasta</th>
        <th align=center class="th2">Observaci&oacute;n</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%
' modificar registros existentes ---------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaestructuras.tenro, evaestructuras.estrnro,evaestructuras.fdesde,evaestructuras.fhasta,"
l_sql = l_sql & " tipoestructura.tedabr, evaluaestr.evaestrdext, estructura.estrdabr"
l_sql = l_sql & " FROM evaestructuras "
l_sql = l_sql & " INNER JOIN tipoestructura ON tipoestructura.tenro = evaestructuras.tenro"
l_sql = l_sql & " INNER JOIN evaluaestr ON evaluaestr.tenro = evaestructuras.tenro "
l_sql = l_sql & "        AND evaluaestr.evldrnro =" & l_evldrnro
l_sql = l_sql & " LEFT JOIN estructura  ON estructura.estrnro = evaestructuras.estrnro"
l_sql = l_sql & " INNER JOIN evaseccestr ON evaseccestr.tenro = evaestructuras.tenro"
l_sql = l_sql & " WHERE evaluaestr.evldrnro =" & l_evldrnro
l_sql = l_sql & "   AND evaestructuras.evacabnro =" & l_evacabnro
l_sql = l_sql & " ORDER BY evaseccestr.orden" 
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("tenro")%>)">
        <td align=center>
			<input type="text" class="rev" name="tedabr<%=l_rs("tenro")%>" size=20 value="<%=l_rs("tedabr")%>" >
		</td>
        <td align=center>
		    <iframe name="ifrm<%= l_rs("tenro")%>" scrolling="No" src="filtroNivel.asp?tipo=<%=l_rs("tenro")%>&estrnro=<%=l_rs("estrnro")%>" width="155" height="25"></iframe> 
		</td>
        <td align=center>
			<input type="text" name="fdesde<%=l_rs("tenro")%>" size="10" maxlength="10" value="<%=l_rs("fdesde")%>">
			<a href="Javascript:Ayuda_Fecha(document.datos.fdesde<%=l_rs("tenro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
        <td align=center>
			<input type="text" name="fhasta<%=l_rs("tenro")%>" size="10" maxlength="10" value="<%=l_rs("fhasta")%>">
			<a href="Javascript:Ayuda_Fecha(document.datos.fhasta<%=l_rs("tenro")%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
        <td align=center>
			<textarea name="evaestrdext<%=l_rs("tenro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaestrdext"))%></textarea>
		</td>
        <td valign=top>
   			<a href=# onclick="if (ValidarDatos(document.ifrm<%=l_rs("tenro")%>,document.datos.fdesde<%=l_rs("tenro")%>,document.datos.fhasta<%=l_rs("tenro")%>,document.datos.evaestrdext<%=l_rs("tenro")%>)) { grabar.location='grabar_estructuras_eva_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&tenro=<%=l_rs("tenro")%>&estrnro='+document.ifrm<%=l_rs("tenro")%>.registro.grupo.value+'&fdesde='+document.datos.fdesde<%=l_rs("tenro")%>.value+'&fhasta='+document.datos.fhasta<%=l_rs("tenro")%>.value+'&evaestrdext='+escape(document.datos.evaestrdext<%=l_rs("tenro")%>.value);document.datos.grabado<%=l_rs("tenro")%>.value='M'; }">Modificar</a>
			<br>
			<a href=# onclick="grabar.location='grabar_estructuras_eva_00.asp?tipo=B&tenro=<%=l_rs("tenro")%>&evldrnro=<%=l_evldrnro%>';document.datos.grabado<%=l_rs("tenro")%>.value='B';">Baja</a>
			<br>
			<input type="text" readonly disabled name="grabado<%=l_rs("tenro")%>" size="1">
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
