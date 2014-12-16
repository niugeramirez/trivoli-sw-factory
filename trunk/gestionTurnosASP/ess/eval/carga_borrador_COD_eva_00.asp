<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_borrador_COD_eva_00.asp
'Objetivo : ABM de compromisos
'Fecha	  : 01-02-2005
'Autor	  : CCRossi
'Modificacion: 18-02-2005 CCRossi - Inicializar con objetivos iniciales si es que NO tiene objetivos ya definidos.
'Modificacion: 18-03-2005 CCRossi - No permitir superar el 100%
'Modificacion: 04-04-2005 CCRossi - Agregar al orden por tipo de comproiso, orden por nro de objetivo.
'            Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
on error goto 0
 Dim l_rs
 Dim l_rs1
 Dim l_cm
 Dim l_sql
 Dim l_filtro
 Dim l_orden

 Dim col
 col=5
 
'locales
 dim l_puntacion
 dim l_puntajemanual
 dim l_puntaje
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evaevenro

 dim l_cabaprobada
 
 dim l_evatipobjnro ' tipo objetivo
 dim l_cantidad ' cantidad de objetivos de un tipo 
 
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_empleg

 dim l_tieneobj
 
'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 
 l_evldrnro = request.querystring("evldrnro")
 l_evapernro = request.querystring("evapernro")

 if l_orden = "" then
  l_orden = " ORDER BY evaobjborrnro "
 end if

'tomar el lolgeado, si hay un logeado viene de AUTOGESTION!
'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaevenro, evatevnro, empleado.empleg, evacab.cabaprobada FROM  evadetevldor INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaluador = l_rs("empleg")
	l_evaevenro = l_rs("evaevenro")
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
 l_sql = "SELECT evacabnro, evatevnro  FROM  evadetevldor WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
 end if
 l_rs.close
 set l_rs=nothing

'buscar puntaje cargado
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT puntaje, puntajemanual, tieneobj FROM  evacab WHERE evacabnro   = " & l_evacabnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_puntajemanual= l_rs("puntajemanual")
	l_puntaje= l_rs("puntaje")
	l_tieneobj= l_rs("tieneobj")
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
l_sql = "SELECT evaobjborr.evaobjborrnro, evaobjborrdext, evldrnro, evaobjborrfmed,evaobjborrpond, evaobjborr.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor FROM evaobjborr INNER JOIN evaluaobjborr ON evaluaobjborr.evaobjborrnro = evaobjborr.evaobjborrnro"
l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjborr.evatipobjnro LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro AND evatipoobjpor.evaevenro = " & l_evaevenro & " WHERE evaluaobjborr.evldrnro =" & l_evldrnro & " ORDER BY evatipoobj.evatipobjorden "
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

if cint(l_haycompromisos)=0 and cint(l_tieneobj)=-1 and cint(l_cabaprobada)=0 then

'Buscar el ternro del supervisor del empleado actual
 Dim l_ternro
 Dim l_evaobjborrnro
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaluador FROM  evadetevldor INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro WHERE evadetevldor.evacabnro   = " & l_evacabnro & " AND evatevnro  = " & cevaluador
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_ternro = l_rs("evaluador")
 end if
 l_rs.close
 set l_rs=nothing

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT DISTINCT evaobjininro,evaobjdext, evaobjformed, evaobjinicial.evatipobjnro , evatipobjdabr FROM  evaobjinicial  INNER JOIN evatipoobj ON evatipoobj.evatipobjnro=evaobjinicial.evatipobjnro WHERE  evaobjinicial.ternro = " & l_ternro & " AND    evaobjinicial.evaevenro = " & l_evaevenro
 rsOpen l_rs, cn, l_sql, 0
 do until l_rs.eof
	' agregar compromiso inicial...
	
	l_sql= "INSERT INTO evaobjborr (evaobjborrdext,evaobjborrfmed,evatipobjnro) "
	l_sql = l_sql & " values ('" & trim(l_rs("evaobjdext")) & "','" & trim(l_rs("evaobjformed")) & "'," & l_rs("evatipobjnro") &")"
	'Response.Write l_sql & "<br>"
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_sql = fsql_seqvalue("evaobjborrnro","evaobjborr")
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	rsOpen l_rs1, cn, l_sql, 0
	'Response.Write l_sql
	if not l_rs1.eof then
		l_evaobjborrnro=l_rs1("evaobjborrnro")
	end if	
	l_rs1.Close
	Set l_rs1 = Nothing
	
	if trim(l_evaobjborrnro)<>"" and not isnull(l_evaobjborrnro) then
		l_sql = "insert into evaluaobjborr (evldrnro,evaobjborrnro) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_evaobjborrnro &")"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		' buscar el evaobjborrnro
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaobjborrnro FROM  evaobjborr WHERE  evaobjborrdext ='" & trim(l_rs("evaobjdext")) &"'"
		l_sql = l_sql & " AND    evaobjborrfmed ='" & trim(l_rs("evaobjformed")) &"'"
		l_sql = l_sql & " AND    evatipobjnro = " & l_rs("evatipobjnro")
		l_sql = l_sql & " ORDER BY evaobjborrnro DESC"
		rsOpen l_rs1, cn, l_sql, 0
 		if not l_rs1.eof then
 			l_sql = "insert into evaluaobjborr (evldrnro,evaobjborrnro) "
			l_sql = l_sql & " values (" & l_evldrnro & "," & l_rs1("evaobjborrnro") &")"
			set l_cm = Server.CreateObject("ADODB.Command")  
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
 		end if
 		l_rs1.close
 		set l_rs1=nothing	
	end if
		
	l_rs.MoveNext
 loop
 l_rs.close
 set l_rs=nothing


end if
'====================================================================================
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

function Controlar(objetivo,ponde,tipo){

	objetivo.value=Blanquear(objetivo.value);
	if (tipo.value=="")
	{
		alert('Seleccione un Tipo de Compromiso.');
		tipo.focus();
		return false;
	}	
	else
	if (objetivo.value.trim()=="")
	{
		alert('Ingrese un Compromiso.');
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
		<%if l_haycompromisos=1 then%>
			if ( Sumar() )
				return true;
			else {
				ponde.focus();
				return false;					
				}
		<%else%>		
			return true;
		<%end if%>	
	}
}	

function Sumar(){

	var formElements = document.datos.elements;
	var total1;
	var indice;
	var nuevotipo;
	var tipobj;
	var cantidad;
	
	var error=0;
	var k=0;
	cantidad=0;
	
	<%
	dim l_tipos
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evaobjborr.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor FROM evaobjborr INNER JOIN evaluaobjborr ON evaluaobjborr.evaobjborrnro = evaobjborr.evaobjborrnro"
	l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjborr.evatipobjnro LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro AND evatipoobjpor.evaevenro = " & l_evaevenro
	l_sql = l_sql & " WHERE evaluaobjborr.evldrnro =" & l_evldrnro & " ORDER BY evatipoobj.evatipobjorden "
	rsOpen l_rs1, cn, l_sql, 0 
	l_tipos = 0
	do until l_rs1.eof%>
		
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
			if (total1 > 100)
				error=1;
			else	
				if (total1 < 100) 
					error=2;

			
		indice = cantidad * 7;
		
		
		document.datos.totalponderacion<%=l_rs1("evatipobjnro")%>.value=total1;
		<%l_tipos= l_tipos + 1%>
		k=indice + <%=l_tipos%>;
		<%l_rs1.MoveNext
	loop
	l_rs1.close	
	set l_rs1=nothing%>
	
	if (error==1) {
		alert('ERROR. El total de Ponderación de cada Tipo de Compromiso debe ser 100.');
		return false;
		}
	else	
		if (error==2) {
			alert('AVISO. El total de Ponderación de cada Tipo de Compromiso debe ser 100.');
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
        <th class="th2">&nbsp;</th>
		<th align=center class="th2">Compromisos</th>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <th align=center class="th2">Ponderaci&oacute;n</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaobjborr.evaobjborrnro, evaobjborrdext, evldrnro, evaobjborrfmed,evaobjborrpond, evaobjborr.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor FROM evaobjborr INNER JOIN evaluaobjborr ON evaluaobjborr.evaobjborrnro = evaobjborr.evaobjborrnro"
l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjborr.evatipobjnro LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro AND evatipoobjpor.evaevenro = " & l_evaevenro & " WHERE evaluaobjborr.evldrnro =" & l_evldrnro & " ORDER BY evatipoobj.evatipobjorden, evaobjborr.evaobjborrnro "
'Response.Write l_sql
rsOpen l_rs, cn, l_sql, 0 
l_evatipobjnro=""
l_cantidad = 0
'Response.Write l_evatipobjnro & "<br>"

do until l_rs.eof
 'Response.write l_rs("evatipobjnro") & "<br>"

	if l_evatipobjnro <> l_rs("evatipobjnro") then
		l_evatipobjnro= l_rs("evatipobjnro") 
		l_cantidad = 1%>
		<tr>
        <td colspan="<%=col%>"><b><%=l_rs("evatipobjdabr")%>&nbsp;<%=l_rs("evatipopor")%>%</b>
		</td>
		</tr>
	<%else
		l_cantidad = l_cantidad + 1
	end if%>
    <tr>
        <td align=center valign=top style="width:10">
			<b><%=l_cantidad%></b>
		</td>
        <td align=center>
			<input type=hidden name="evatipobjnro<%=l_rs("evaobjborrnro")%>" value="<%=l_rs("evatipobjnro")%>">
        	<textarea name="evaobjborrdext<%=l_rs("evaobjborrnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjborrdext"))%></textarea>
		</td>
        <td align=center>
			<textarea name="evaobjborrfmed<%=l_rs("evaobjborrnro")%>"  maxlength=200 size=200 cols=30 rows=4><%=trim(l_rs("evaobjborrfmed"))%></textarea>
		</td>

        <td align=center>
			<input pond="<%=l_rs("evaobjborrnro")%>" type="text" name="evaobjborrpond<%=l_rs("evaobjborrnro")%>" size=5 maxlength=3 value="<%=l_rs("evaobjborrpond")%>" >
			<input class="blanc" type="hidden" name="evaobjalcanza<%=l_rs("evaobjborrnro")%>" size=5>
			<input readonly class="blanc" type="hidden" name="puntacion<%=l_rs("evaobjborrnro")%>" size=5 value="<%=l_puntacion%>">
		</td>
        <td valign=top>
			<a href=# onclick="if (Controlar(document.datos.evaobjborrdext<%=l_rs("evaobjborrnro")%>,document.datos.evaobjborrpond<%=l_rs("evaobjborrnro")%>,'tipo')) { if(ValidarDatos(document.datos.evaobjborrpond<%=l_rs("evaobjborrnro")%>)) {grabar.location='grabar_borrador_COD_eva_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&evaobjborrnro=<%=l_rs("evaobjborrnro")%>&evaobjborrdext='+escape(document.datos.evaobjborrdext<%=l_rs("evaobjborrnro")%>.value)+'&evaobjborrfmed='+escape(Blanquear(document.datos.evaobjborrfmed<%=l_rs("evaobjborrnro")%>.value))+'&evaobjborrpond='+document.datos.evaobjborrpond<%=l_rs("evaobjborrnro")%>.value+'&evatipobjnro='+document.datos.evatipobjnro<%=l_rs("evaobjborrnro")%>.value;document.datos.grabado<%=l_rs("evaobjborrnro")%>.value='M';} } else document.datos.grabado<%=l_rs("evaobjborrnro")%>.value='';">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evaobjborrnro")%>" size="1">
			<br>
			<a href=# style="color:red;" onclick="if (confirm('¿ Desea Eliminar el Compromiso?')==true) { grabar.location='grabar_borrador_COD_eva_00.asp?tipo=B&evaobjborrnro=<%=l_rs("evaobjborrnro")%>&evldrnro=<%=l_evldrnro%>&evaobjborrdext='+escape(Blanquear(document.datos.evaobjborrdext<%=l_rs("evaobjborrnro")%>.value))+'&evaobjborrfmed='+escape(Blanquear(document.datos.evaobjborrfmed<%=l_rs("evaobjborrnro")%>.value))+'&evaobjborrpond='+document.datos.evaobjborrpond<%=l_rs("evaobjborrnro")%>.value};document.datos.grabado<%=l_rs("evaobjborrnro")%>.value='B';">Eliminar Compromiso</a>
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
		</td>
		<td>&nbsp;</td>
		</tr>
		<%
	end if
loop
l_rs.Close
set l_rs = Nothing
%>
	<tr>
		<td colspan="<%=col%>" align=left><b>Tipo de Compromiso:</B>
			<%' BUSCAR tipo objetivos
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evatipobjnro,evatipobjdabr FROM evatipoobj "
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
    <tr>
		<td>&nbsp;</td>
        <td align=center >
			<textarea name="evaobjborrdext"  maxlength=200 size=200 cols=40 rows=4></textarea>
		</td>
		<td align=center>
			<textarea name="evaobjborrfmed"  maxlength=200 size=200 cols=30 rows=4></textarea>
		</td>
        <td align=center >
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input pond="pond"  type="text" name="evaobjborrpond" size=5 maxlength=3>
			<input class="blanc" type="text" name="evaobjalcanza" size=5>
			<input readonly type="text" class="blanc" name="puntacion" size=5>
		</td>
		<td valign=top >
			<a href=# onclick="javascript:if (Controlar(document.datos.evaobjborrdext,document.datos.evaobjborrpond,document.datos.evatipobjnro)) { if (ValidarDatos(document.datos.evaobjborrpond)) {grabar.location='grabar_borrador_COD_eva_00.asp?tipo=A&evldrnro=<%=l_evldrnro%>&evaobjborrdext='+escape(Blanquear(document.datos.evaobjborrdext.value))+'&evaobjborrpond='+document.datos.evaobjborrpond.value+'&evatipobjnro='+document.datos.evatipobjnro.value+'&evaobjborrfmed='+escape(Blanquear(document.datos.evaobjborrfmed.value));document.datos.grabado.value='G'; } }">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
		</td>
    </tr>
    
</table>
<input type="Hidden" name="cabnro" value="0">
<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar"-->

<%
cn.Close
set cn = Nothing
%>
</form>
</body>
</html>
