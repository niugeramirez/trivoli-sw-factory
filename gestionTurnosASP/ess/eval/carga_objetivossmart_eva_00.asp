<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_objetivossmart_eva_00.asp
'Objetivo : ABM de objetivos smart de evaluacion
'Fecha	  : 14-05-2004
'Autor	  : CCRossi
'Modificacion: 28-07-2004 CCRossi cambiar el tema del puntaje. ahora unico de evacab.
'Modificacion: 22-10-2004 Agegar cantidad de obj de cada tipo
'Modificacion: 22-10-2004 Cambiar "MOdificar" por "Grabar"
'Modificacion: 29-10-2004 CCROssi- Controlar que sea readonly para el no loogeado 
'			desde autogestion
'Modificacion: 19-11-04-CCRossi-  control de caraceteres raros
'Modificacion: 27-12-04-CCRossi-  cambiar nombre de link BAJA
'Modificacion: 03-02-2005-CCRossi-  pequeñas adecuaciones codelco.
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'            26-09-2006 - Gustavo Ring - Control de cantidad máxima de caracteres  
'			  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 Dim l_rs
 Dim l_sql
 Dim l_filtro
 Dim l_orden

 Dim col
 if cformed = -1 then
	col=5
 else 
	col=4	
 end if	

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
 l_sql = "SELECT evaevenro, evatevnro, empleado.empleg  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaluador = l_rs("empleg")
	l_evaevenro = l_rs("evaevenro")
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
 l_sql = "SELECT puntaje, puntajemanual  "
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
	
	if (objetivo.value.length > 300) 
		{
    	alert("Los Objetivos no deben superar los 300 caracteres.");
		document.datos.habdesext.focus();
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
		var formElements = document.datos.elements;
		
		var total=0;
		
		// controlar PONDERACION
		var i=3;
		var total1=0;
		while (i<formElements.length-2)
		{
			total1 = Number(total1) + Number(formElements[i].value);
			i = i+7;
		}
		document.datos.totalponderacion.value=total1;

		
		if ((Number(total1)!==100))
		{
			alert('El total de Ponderación debe ser 100.');
			if ((Number(total1) < 100) && (total1!==""))
				return true;
			else
			{
				ponde.focus();
				return false;
			}	
		}
		else
			return true;	
	}
}	

function Sumar(){
	var formElements = document.datos.elements;
	
	var total=0;
	var i=3;
	while (i<formElements.length-3)
	{
		total = Number(total) + Number(formElements[i].value);
		i = i+7;
	}
	document.datos.totalponderacion.value=total;
		
	if ((total > 100) || (total < 100)) {
		alert('El total de Ponderación debe ser 100.');
		return false;
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
        <td colspan="<%=col + 1%>"><b><%=l_rs("evatipobjdabr")%>
        <%if ccodelco=-1 then%>
			&nbsp;<%=l_rs("evatipopor")%>%
		<%end if%>	
        </b>
		</td>
		</tr>
	<%else
		l_cantidad = l_cantidad + 1
	end if%>
    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("evaobjnro")%>)">
        <td align=center valign=top style="width:10">
			<b><%=l_cantidad%></b>
		</td>
        <td align=center>
			<input type=hidden name="evatipobjnro<%=l_rs("evaobjnro")%>" value="<%=l_rs("evatipobjnro")%>">
        	<textarea name="evaobjdext<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjdext"))%></textarea>
		</td>
		<%if cformed=-1 then%>
        <td align=center>
			<textarea name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjformed"))%></textarea>
		</td>
		<%else%>
		<input type="hidden" name="evaobjformed<%=l_rs("evaobjnro")%>" value="<%=l_rs("evaobjformed")%>">
		<%end if%>
        <td align=center>
			<input pond="<%=l_rs("evaobjnro")%>" type="text" name="evaobjpond<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjpond")%>" >
			<input type="hidden" name="evaobjalcanza<%=l_rs("evaobjnro")%>" size=5 value="<%=l_rs("evaobjalcanza")%>">
			<input readonly type="hidden" class="blanc" name="puntacion<%=l_rs("evaobjnro")%>" size=5 value="<%=l_puntacion%>">
		</td>
        <td valign=top>
			<a href=# onclick="if (Controlar(document.datos.evaobjdext<%=l_rs("evaobjnro")%>,document.datos.evaobjpond<%=l_rs("evaobjnro")%>,document.datos.evaobjalcanza<%=l_rs("evaobjnro")%>,'tipo')) { if(ValidarDatos(document.datos.evaobjpond<%=l_rs("evaobjnro")%>)) {grabar.location='grabar_objetivossmart_eva_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evapernro=<%=l_evapernro%>&evaobjnro=<%=l_rs("evaobjnro")%>&evaobjdext='+escape(document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value)+'&evaobjformed='+escape(Blanquear(document.datos.evaobjformed<%=l_rs("evaobjnro")%>.value))+'&evaobjpond='+document.datos.evaobjpond<%=l_rs("evaobjnro")%>.value+'&evaobjalcanza='+document.datos.evaobjalcanza<%=l_rs("evaobjnro")%>.value;document.datos.grabado<%=l_rs("evaobjnro")%>.value='M';}}">Grabar</a>
			<br>
			<a href=# onclick="grabar.location='grabar_objetivossmart_eva_00.asp?tipo=B&evaobjnro=<%=l_rs("evaobjnro")%>&evapernro=<%=l_evapernro%>&evldrnro=<%=l_evldrnro%>&evaobjdext='+escape(Blanquear(document.datos.evaobjdext<%=l_rs("evaobjnro")%>.value))+'&evaobjformed='+escape(Blanquear(document.datos.evaobjformed<%=l_rs("evaobjnro")%>.value))+'&evaobjpond='+document.datos.evaobjpond<%=l_rs("evaobjnro")%>.value+'&evaobjalcanza='+document.datos.evaobjalcanza<%=l_rs("evaobjnro")%>.value;document.datos.grabado<%=l_rs("evaobjnro")%>.value='B';">Eliminar <%if ccodelco=-1 then%>Compromiso<%else%>Objetivo<%end if%></a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evaobjnro")%>" size="1">
		</td>
    </tr>
<%
	l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
%>
	<tr>
	<td>&nbsp;</td>
	<td colspan=5 align=left><b>Tipo del Nuevo Objetivo:</B>
			<%' BUSCAR tipo objetivos
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evatipobjnro,evatipobjdabr "
			l_sql = l_sql & "FROM evatipoobj "
			rsOpen l_rs, cn, l_sql, 0 %>
			<select name="evatipobjnro">
			<%do while not l_rs.eof%>
				<option value=<%=l_rs("evatipobjnro")%>><%=l_rs("evatipobjdabr")%></option>
			<%l_rs.MoveNext
			loop
			l_rs.Close
			set l_rs = nothing%>
			</select>
		</td>
	</tr>	
    <tr onclick="Javascript:Seleccionar(this,0)">
		<td>&nbsp;</td>
        <td align=center >
			<textarea name="evaobjdext"  maxlength=200 size=200 cols=30 rows=4></textarea>
		</td>
		<%if cformed=-1 then%>
        <td align=center>
			<textarea name="evaobjformed"  maxlength=200 size=200 cols=30 rows=4></textarea>
		</td>
		<%else%>
		<input type="hidden" name="evaobjformed">
		<%end if%>
        <td align=center >
			<input pond="pond"  type="text" name="evaobjpond" size=5>
			<input type="hidden" name="evaobjalcanza" size=5>
			<input readonly type="hidden" class="blanc" name="puntacion" size=5>
		</td>
		<td valign=top >
			<a href=# onclick="javascript:if (Controlar(document.datos.evaobjdext,document.datos.evaobjpond,document.datos.evaobjalcanza,document.datos.evatipobjnro)) { if (ValidarDatos(document.datos.evaobjpond)) {grabar.location='grabar_objetivossmart_eva_00.asp?tipo=A&evapernro=<%=l_evapernro%>&evldrnro=<%=l_evldrnro%>&evaobjdext='+escape(Blanquear(document.datos.evaobjdext.value))+'&evaobjformed='+escape(Blanquear(document.datos.evaobjformed.value))+'&evaobjpond='+document.datos.evaobjpond.value+'&evaobjalcanza='+document.datos.evaobjalcanza.value+'&evatipobjnro='+document.datos.evatipobjnro.value;document.datos.grabado.value='G'; } }">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
		</td>
    </tr>
	<!-- t o t a  l e s ----------------------------------->
    <tr>
		<%if cformed=-1 then%>
			<td align=center>&nbsp;	</td>
		<%end if%>
        <td>&nbsp;</td>
        <td align=right><b>Total</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion" size=5>
			<input style="background : #e0e0de;" readonly class="blanc" type="hidden" name="totalpuntuacion" size=5>
		</td>
		<td>&nbsp;</td>
    </tr>
    
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
