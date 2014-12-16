<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : ver_compromisos_COD_eva_00.asp
'Objetivo : Ver de compromisos.sumatoria de 100 por cada tipo
'Fecha	  : 07-02-2005
'Autor	  : CCRossi
'Modificacion: 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 on error goto 0
 Dim l_rs
 Dim l_cm
 Dim l_rs1
 Dim l_sql
 Dim l_filtro
 Dim l_orden

 Dim col
 if cformed = -1 then
	col=5
 else 
	col=4	
 end if	

 dim l_cabaprobada
 dim l_tieneobj
 
'locales
 dim l_puntacion
 dim l_puntajemanual
 dim l_puntaje
 dim l_evacabnro 
 dim l_evatevnro 

 dim l_evatipobjnro ' tipo objetivo
 dim l_cantidad ' cantidad de objetivos de un tipo 
 
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_empleg

 dim l_evaevenro ' necesario para buscar porcentajes de tipo objetivos para CODELCO
'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
 end if

' tomar el lolgeado, si hay un logeado viene de AUTOGESTION!
'buscar la evaluador y evaevenro
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaevenro, evatevnro, empleado.empleg, tieneobj, cabaprobada  FROM  evadetevldor INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaluador = l_rs("empleg")
	l_evaevenro = l_rs("evaevenro")
	l_tieneobj  = l_rs("tieneobj")
	l_cabaprobada = l_rs("cabaprobada")
 end if
 l_rs.close
 set l_rs=nothing
 
 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	
 dim l_mostrar 
 l_mostrar = "0"

'Response.Write l_empleg & "<br>" & l_evaluador
 if trim(l_empleg)<>"" and not isNull(l_empleg) then
	if trim(l_empleg) = trim(l_evaluador) then
		l_mostrar = "1"
	else	
		l_mostrar = "0"
	end if
 else
 	l_mostrar = "1"
 end if 
 
'___________________________________________________________________________________
function PasarComaAPunto(valor)
	dim l_numero
	dim l_ubicacion
	dim l_entero
	dim l_decimal
	l_numero = trim(valor)
	l_ubicacion = InStr(l_numero, ",")
	if l_ubicacion > 1 then
		l_ubicacion = l_ubicacion  - 1
		l_entero = left(l_numero, l_ubicacion)
		l_ubicacion = l_ubicacion  + 1
		l_decimal = right(l_numero, (len(l_numero) - l_ubicacion))
    	l_numero = l_entero & "." & l_decimal
    	PasarComaAPunto = l_numero
    else
		PasarComaAPunto = valor
	end if
end function	

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

'buscar puntaje cargado
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT puntaje, puntajemanual "
 l_sql = l_sql & " FROM  evacab "
 l_sql = l_sql & " WHERE evacabnro   = " & l_evacabnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_puntajemanual= l_rs("puntajemanual")
	l_puntaje= l_rs("puntaje")
 end if
 l_rs.close
 set l_rs=nothing
 
 if trim(l_puntajemanual)<>"" then
	l_puntajemanual = PasarComaAPunto(l_puntajemanual)
 else	
	l_puntajemanual = ""
 end if	
 
 
 '===================================================================================
' Chequear si hay o no objetivos para realizar la Suma en Javascript
dim l_haycompromisos
l_haycompromisos=0
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evldrnro, evaobjpond, evaobjalcanza,evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden ,evatipopor FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro AND evaluaobj.evaborrador = 0  LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro  LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro AND evatipoobjpor.evaevenro = " & l_evaevenro
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro & " ORDER BY evatipoobj.evatipobjorden "
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_haycompromisos=1
end if
l_rs.close
set l_rs=nothing	

'Response.Write l_haycompromisos & "<br>"
'Response.Write l_tieneobj & "<br>"
'Response.Write l_cabaprobada & "<br>"

 'si no hay compromisos crearlos con los INICIALES =======================================

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

var entrada = new Array;
var salida = new Array;

entrada[0]=50;
entrada[1]=60;
entrada[2]=70;
entrada[3]=80;
entrada[4]=90;
entrada[5]=99;
entrada[6]=110;
entrada[7]=120;
entrada[8]=130;
entrada[9]=140;

salida[0]=0;
salida[1]=0.5;
salida[2]=1;
salida[3]=1.5;
salida[4]=2;
salida[5]=2.5;
salida[6]=3;
salida[7]=3.5;
salida[8]=4;
salida[9]=4.5;
salida[10]=5;

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

function Controlar(objetivo,ponde,alcanza,tipo){

	objetivo.value=Blanquear(objetivo.value);
	if (tipo.value=="")
	{
		alert('Seleccione un Tipo de Objetivo.');
		tipo.focus();
		return false;
	}	
	else
	if (objetivo.value.trim()=="")
	{
		alert('Ingrese un Objetivo.');
		tipo.focus();
		return false;
	}	
	else
	if (ponde.value>100)
	{
		alert('La Ponderación debe ser menor o igual a 100.');
		ponde.focus();
		return false;
	}	
	else
	{	
		Sumar();
		return true;		
	}
}	

function Sumar(){

	var formElements = document.datos.elements;
	var total1=0;
	
	var indice;
	var nuevotipo;
	var k=0;
	var error=0;
	
	var tipobj=formElements[k].value;
	var cantidad=0;
	while (k<formElements.length-7)
	{
		if (formElements[k].value==tipobj)
		{
			total1=Number(total1) + Number(formElements[k+3].value);
			cantidad=cantidad + 1;
			k = k+7;
		}	
		else
			k = k + 99;	// salir del bucle no vale la pena seguir.
	}
	if (total1!==0)
		if ((total1 > 100) || (total1 < 100)) 
			error=1;
		
	indice = cantidad * 7;
	formElements[indice].value=total1;
	
	//--------------------------------****------------------------
	// desde aca hasta marca copiar cuando se agreguen mas tipos de compromisos
	k=indice + 1; // sumar 1 por el input del total
	total1=0;
	tipobj=formElements[k].value;
	while (k<formElements.length-7)
	{
		if (formElements[k].value==tipobj)
		{
			total1=Number(total1) + Number(formElements[k+3].value);
			cantidad=cantidad + 1;
			k = k+7;
		}	
		else
			k = k + 99;	// salir del bucle no vale la pena seguir.
	}
	if (total1!==0)
		if ((total1 > 100) || (total1 < 100)) 
			error=1;
	
	indice = (cantidad * 7) + 1;
	formElements[indice].value=total1;
	//--------------------------------****------------------------
	
	if (error==1) {
		alert('El total de Ponderación de cada Tipo de Compromiso debe ser 100.');
		return true;
		}
	else	
		return true;
	
}	
	

function ValidarDatos(ponde)
{
	if (ponde.value=="") 
	{
		alert('Ingrese una Ponderación.');
		ponde.focus();
		return false;
	}	
	else
	if (isNaN(ponde.value)) 
	{
		alert('Ingrese una Ponderación válida.');
		ponde.focus();
		return false;
	}	
	else
		return true;
		
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="Sumar();">
<table>
    <tr>
		<%if ccodelco=-1 then%>
        <th align=center colspan=2 class="th2">Compromisos</th>
        <%else%>
        <th align=center colspan=2 class="th2">Objetivos</th>
        <%end if%>
        <%if cformed=-1 or ccodelco=-1 then%>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <%end if%>
        <th align=center class="th2">Ponderaci&oacute;n</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, evldrnro, evaobjpond, evaobjalcanza,"
l_sql = l_sql & " evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden "
if ccodelco=-1 then
l_sql = l_sql & " ,evatipopor "
end if
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
l_sql = l_sql & "		 AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
if ccodelco=-1 then
l_sql = l_sql & " LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro"
l_sql = l_sql & "		 AND evatipoobjpor.evaevenro = " & l_evaevenro
end if
l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
l_evatipobjnro=""
l_cantidad = 0
do until l_rs.eof
	if l_evatipobjnro <> l_rs("evatipobjnro") then
		l_evatipobjnro= l_rs("evatipobjnro") 
		l_cantidad = 1%>
		<tr>
			<td colspan="<%=col + 1%>"><b><%=l_rs("evatipobjdabr")%>&nbsp;<%=l_rs("evatipopor")%>%</b></td>
		</tr>
	<%else
		l_cantidad = l_cantidad + 1
	end if%>
    <tr>
        <td align=center valign=top style="width:10">
			<b><%=l_cantidad%></b>
		</td>
        <td align=center>
			<input type=hidden name="evatipobjnro<%=l_rs("evaobjnro")%>" value="<%=l_rs("evatipobjnro")%>">
        	<textarea readonly style="background : #e0e0de;" name="evaobjdext<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
		<td align=center>
			<textarea readonly style="background : #e0e0de;" name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjformed"))%></textarea>
		</td>
		<td align=center>
			<input readonly style="background : #e0e0de;" pond="<%=l_rs("evaobjnro")%>" type="text" name="evaobjpond<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjpond")%>" >
			<input type="hidden" name="evaobjalcanza<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjalcanza")%>">
			<input readonly type="text" class="blanc" name="puntacion<%=l_rs("evaobjnro")%>" size=5 value="<%=l_puntacion%>">
		</td>
        <td valign=top>
			<input class="rev" type="hidden" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evaobjnro")%>" size="1">
		</td>
    </tr>
<%
	l_rs.MoveNext
	if not l_rs.eof then
		if l_rs("evatipobjnro") <> l_evatipobjnro then%>
		<!-- t o t a  l e s ----------------------------------->
		<tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td align=right><b>Total</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion<%=l_evatipobjnro%>" size=5>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
		<td>&nbsp;</td>
		</tr>
		<%end if
	else
		%>
		<!-- t o t a  l e s ----------------------------------->
		<tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td align=right><b>Total</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion<%=l_evatipobjnro%>" size=5>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
		<td>&nbsp;</td>
		</tr>
		<%
	end if
loop
l_rs.Close
set l_rs = Nothing
%>
</table>
<input type="Hidden" name="cabnro" value="0">
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">

<%
cn.Close
set cn = Nothing
%>
</form>
</body>
</html>
