<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->

<script>
	var resultados="";
	var notafinal="";
	var correspondencia = new Array;	
	
	correspondencia[0]= "No cumple."
	correspondencia[1]= "No cumple."
	correspondencia[2]= "Cumple en parte."
	correspondencia[3]= "Cumple."
	correspondencia[4]= "Supera lo comprometido."
	correspondencia[5]= "Excelencia."
	
	notafinal = notafinal + "1.0 -- 1.9: No cumple \n\n"
	notafinal = notafinal + "2.0 -- 2.9: Cumple en parte\n\n"
	notafinal = notafinal + "3.0 -- 3.9: Cumple\n\n"
	notafinal = notafinal + "4.0 -- 4.9: Supera lo comprometido\n\n"
	notafinal = notafinal + "       5.0: Excelencia"
	
</script>
<% 
'=====================================================================================
'Archivo  : Ver_evalborrador_COD_eva_00.asp
'Objetivo : Carga Evaluacion Borrador modo READONLY
'Fecha	  : 16-02-2005
'Autor	  : CCRossi
'Modificacion: 21-03-2005 Cambiar tamaño de letra en clase ".rev"
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
on error goto 0

 Dim l_rs
 Dim l_cm
 Dim l_rs1
 Dim l_rs2
 Dim l_sql
 Dim l_filtro
 Dim l_orden

 Dim l_col
 l_col=8

'locales
 dim l_puntaje
 dim l_puntajemanual
 dim l_evacabnro 
 dim l_evatevnro 
 dim l_evatevlogeado
 dim l_ternrologeado
 
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_mostrar '1 o 0 si tiene que mostrar la observacion. 
 dim l_evatipobjnro ' tipo objetivo
 dim l_cantidad ' cantidad de objetivos de un tipo
 dim l_evaevenro

'parametros
 Dim l_evldrnro
 Dim l_evapernro 'periodo de evaluacion
 dim l_empleg
  
 l_evldrnro	= request.querystring("evldrnro")
 l_evapernro= request.querystring("evapernro")

 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	

 if l_orden = "" then
  l_orden = " ORDER BY evaobjnro "
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
'___________________________________________________________________________________


'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evadetevldor.evacabnro, evatevnro, empleado.empleg  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evacabnro = l_rs("evacabnro")
	l_evatevnro = l_rs("evatevnro")
	l_evaluador = l_rs("empleg")
 end if
 l_rs.close
 set l_rs=nothing

'buscar puntaje cargado
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaevenro, puntaje, puntajemanual  "
 l_sql = l_sql & " FROM  evacab "
 l_sql = l_sql & " WHERE evacabnro   = " & l_evacabnro
  rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_puntajemanual= l_rs("puntajemanual")
	l_puntaje	   = l_rs("puntaje")
	l_evaevenro = l_rs("evaevenro")
 end if
 l_rs.close
 set l_rs=nothing

 if trim(l_puntajemanual)<>"" then
	l_puntajemanual = PasarComaAPunto(l_puntajemanual)
 else	
	l_puntajemanual = ""
 end if	
 
' Controlar si el logeado puede ver o no===============================================
 l_mostrar = "0"
' Response.Write l_empleg & "<br>" & l_evaluador
 if trim(l_empleg)<>"" and not isNull(l_empleg) then
	'buscar la el ternro del logeado
	 Set l_rs = Server.CreateObject("ADODB.RecordSet")
	 l_sql = "SELECT ternro FROM  empleado WHERE empleg   = " & l_empleg
	 rsOpen l_rs, cn, l_sql, 0
	 if not l_rs.EOF then
		l_ternrologeado = l_rs("ternro")
	 end if
	 l_rs.close
	 set l_rs=nothing

	'buscar evatevnro del logeado
	 Set l_rs = Server.CreateObject("ADODB.RecordSet")
	 l_sql = "SELECT evatevnro FROM  evacab "
	 l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
	 l_sql = l_sql & " WHERE evacab.evacabnro   = " & l_evacabnro
	 l_sql = l_sql & " AND   evaluador   = " & l_ternrologeado
	  rsOpen l_rs, cn, l_sql, 0
	 if not l_rs.EOF then
		l_evatevlogeado = l_rs("evatevnro")
	 end if
	 l_rs.close
	 set l_rs=nothing
	 
	if (trim(l_empleg) = trim(l_evaluador)) or (l_evatevlogeado= cgarante) then
		l_mostrar = "1"
	else	
		l_mostrar = "0"
	end if
 else
 	l_mostrar = "1"
 end if

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaluaobj.evaobjnro FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro=evaobjetivo.evaobjnro "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evaluaobj.evldrnro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof 
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * FROM evaluaobj "
	l_sql = l_sql & " WHERE evaobjnro = " & l_rs("evaobjnro")
	l_sql = l_sql & " AND   evldrnro  = " & l_evldrnro
	rsOpen l_rs2, cn, l_sql, 0
	if l_rs2.eof then
		l_sql= "insert into evaluaobj (evldrnro,evaobjnro,evaborrador) "
		l_sql = l_sql & " values (" & l_evldrnro & "," & l_rs("evaobjnro") &",0)"
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	else
		l_sql= "UPDATE evaluaobj SET "
		l_sql = l_sql & " evaborrador = 0 "
		l_sql = l_sql & " WHERE evldrnro = " & l_evldrnro 
		l_sql = l_sql & " AND   evaobjnro= " & l_rs("evaobjnro")
		set l_cm = Server.CreateObject("ADODB.Command")  
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if	
	l_rs2.Close
	set l_rs2=nothing
	
	l_rs.MoveNext
loop
l_rs.close
set l_rs=nothing

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evatrnro, evatrvalor, evatrdesabr, evatrrdesext "
l_sql = l_sql & " FROM evatipresu "
l_sql = l_sql & " WHERE evatrtipo=2" 'Objetivos=Compromisos
l_sql = l_sql & " ORDER BY evatrvalor "
rsOpen l_rs1, cn, l_sql, 0 
do until l_rs1.eof%>
	<script>
	resultados = resultados +'<%=l_rs1("evatrvalor")%>'+":  "+'<%=l_rs1("evatrdesabr")%>'+" - "+'<%=l_rs1("evatrrdesext")%>'+" \n\n";
	</script>
<%l_rs1.MoveNext
 loop
 l_rs1.Close
 set l_rs1=nothing%>
 
 
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function ValidarDatos(alcanza)
{
	if (alcanza.value=="") 
	{
		alert('Ingrese una Nota.');
		alcanza.focus();
		return false;
	}	
	else
	if (isNaN(alcanza.value))
	{
		alert('Ingrese una Nota válida.\nUtilice punto (".") para decimales.');
		alcanza.focus();
		return false;
	}	
	else
	if (!validanumero(alcanza, 1, 1)) 
	{
		alert('Ingrese una Nota con 1 entero y 1 decimal, con "." como separador de decimales.');
		alcanza.focus();
		return false;
	}
	else	
	if ((Number(alcanza.value) >5) || (Number(alcanza.value)<1)) 
	{
		alert('Ingrese una Nota entre 1 y 5.');
		alcanza.focus();
		return false;
	}	
	else
		return true;	
		
}

function Controlar(alcanza,texto){
	if (alcanza.value=="") 
	{
		alert('Ingrese una Nota.');
		alcanza.focus();
		return false;
	}	
	else
	if (isNaN(alcanza.value))
	{
		alert('Ingrese una Nota válida.\nUtilice punto (".") para decimales.');
		alcanza.focus();
		return false;
	}	
	else	
	if (!validanumero(alcanza, 1, 1)) 
	{
		alert('Ingrese una Nota con 1 entero y 1 decimal, con "." como separador de decimales.');
		alcanza.focus();
		return false;
	}
	else
	if ((Number(alcanza.value) >5) || (Number(alcanza.value)<1)) 
	{
		alert('Ingrese una Nota entre 1 y 5.');
		alcanza.focus();
		return false;
	}	
	else
	if (texto.value.length>300)
	{
		alert('La Observación no puede superar 300 caracteres.');
		texto.focus();
		return false;
	}	
	else
	{	
		Sumar();
		Puntuacion();
		return true;		
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
	l_sql = "SELECT DISTINCT evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor "
	l_sql = l_sql & " FROM evaobjetivo "
	l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
	l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
	l_sql = l_sql & " LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro"
	l_sql = l_sql & "		 AND evatipoobjpor.evaevenro = " & l_evaevenro
	l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
	l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
	rsOpen l_rs1, cn, l_sql, 0 
	l_tipos= 0
	do until l_rs1.eof%>
		
		total1=0;
		tipobj=formElements[k].value;
		while (k<formElements.length-4)
		{
			if (formElements[k].value==tipobj)
			{
				total1=Number(total1) + Number(formElements[k+2].value);
				cantidad=cantidad + 1;
				k = k+7;
			}	
			else
				k = k + 99;	// salir del bucle no vale la pena seguir.
		}
		if (total1!==0)
			if ((total1 > 100) || (total1 < 100)) 
				error=1;
			
		document.datos.totalponderacion<%=l_rs1("evatipobjnro")%>.value= total1;
		indice = (cantidad * 7);
		<%l_tipos= l_tipos + 2%>
		k=indice  + <%=l_tipos%>;
		//alert(k);
		<%		
		l_rs1.MoveNext
	loop
	l_rs1.close	
	set l_rs1=nothing%>
	
	if (error==1) {
		alert('El total de Ponderación de cada Tipo de Compromiso debe ser 100.');
		return true;
		}
	else	
		return true;
}	

function Puntuacion(){

	var formElements = document.datos.elements;
	var total1;
	var indice;
	var nuevotipo;
	var tipobj;
	var cantidad;
	
	var error=0;
	var k=0;
	cantidad=0;
	document.datos.notafinalponderada.value=0;
	
	<%
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor "
	l_sql = l_sql & " FROM evaobjetivo "
	l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro"
	l_sql = l_sql & " LEFT  JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
	l_sql = l_sql & " LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro"
	l_sql = l_sql & "		 AND evatipoobjpor.evaevenro = " & l_evaevenro
	l_sql = l_sql & " WHERE evaluaobj.evldrnro =" & l_evldrnro
	l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
	rsOpen l_rs1, cn, l_sql, 0 
	l_tipos=0
	
	do until l_rs1.eof%>
		
		total1=0;
		tipobj=formElements[k].value;
		
		while (k<formElements.length-1)
		{
			if (formElements[k].value==tipobj)
			{
				formElements[k+4].value= formElements[k+2].value * formElements[k+3].value / 100;
				total1 = Number(total1) + Number(formElements[k+4].value); 
				cantidad=cantidad + 1;
				k = k+7;
			}	
			else
				k = k + 99;	// salir del bucle no vale la pena seguir.
		}
		/*if (total1!==0)
			if ((total1 > 100) || (total1 < 100)) 
				error=1; */
			
		document.datos.nota<%=l_rs1("evatipobjnro")%>.value=FormatNumber(total1,2,true,false,false);
			
		// Grabar la nota del tipo ded objetivo (ambito) en EvaPuntaje
		//abrirVentanaH('grabar_nota_COD_eva_00.asp?evldrnro=<%=l_evldrnro%>&evatipobjnro=<%=l_rs1("evatipobjnro")%>&nota='+document.datos.nota<%=l_rs1("evatipobjnro")%>.value,'',5,5); 
		
		correspondencia<%=l_rs1("evatipobjnro")%>.style.visibility='VISIBLE';
		correspondencia<%=l_rs1("evatipobjnro")%>.style.lineHeight='100%';
		correspondencia<%=l_rs1("evatipobjnro")%>.innerHTML="<b>"+correspondencia[parseInt(total1)]+"</b>";
	
		document.datos.notafinalponderada.value=Number(document.datos.notafinalponderada.value) + total1 * <%=l_rs1("evatipopor")%> / 100
		document.datos.notafinalponderada.value=FormatNumber(document.datos.notafinalponderada.value,2,true,false,false);
		
		//abrirVentanaH('grabar_puntuacion_eva_00.asp?evldrnro=<%=l_evldrnro%>&puntaje='+document.datos.notafinalponderada.value,'',5,5);
		
		correspondencianotafinal.style.visibility='VISIBLE';
		correspondencianotafinal.style.lineHeight='100%';
		correspondencianotafinal.innerHTML="<b>"+correspondencia[parseInt(document.datos.notafinalponderada.value)]+"</b>";
		
		//alert(Math.trunc(total1));
		
		indice = (cantidad * 7);
		<%l_tipos= l_tipos + 2%>
		k = indice  + <%=l_tipos%>;
		//alert(k);
		<%		
		l_rs1.MoveNext
	loop
	l_rs1.close	
	set l_rs1=nothing%>
	
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
	font-size: 11;
	border-style: none;
	background : transparent;
}
.total
{
	font-size: 12;
	FONT-WEIGHT: bold;
	background : transparent;
}
.rev
{
	font-size: 11;
	border-style: none;
}
</style>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="<%if l_mostrar<>0 then%>Sumar();Puntuacion();<%end if%>">
<table>
    <tr>
		<th class="th2">&nbsp;</th>
        <th align=center class="th2">Compromiso</th>
        <th align=center class="th2">Forma de Medici&oacute;n</th>
        <th align=center class="th2">Ponderaci&oacute;n</th>
        <th align=center class="th2">Nota</th>
        <th align=center class="th2">Nota Ponderada</th>
        <th align=center class="th2">Observaci&oacute;n</th>
        <th class="th2">&nbsp;</th>
    </tr>
<form name="datos" method="post">
<%
if l_mostrar=0 then%>
	<tr>
        <th align=center colspan="<%=l_col%>" class="th2"> - NO HABILITADO -</th>
    </tr>
<%else

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjdext, evaobjformed, evaobjetivo.evaobjpond, "
l_sql = l_sql & " evaluaobj.evldrnro, evaluaobj.evaobjalcanza, "
l_sql = l_sql & " evaobjsgto.evasgtotexto, "
l_sql = l_sql & " evaobjetivo.evatipobjnro, evatipobjdabr ,evatipobjorden, evatipopor "
l_sql = l_sql & " FROM evaobjetivo "
l_sql = l_sql & " INNER JOIN evaluaobj  ON evaluaobj.evaobjnro	   = evaobjetivo.evaobjnro"
l_sql = l_sql & " INNER JOIN evatipoobj ON evatipoobj.evatipobjnro = evaobjetivo.evatipobjnro"
l_sql = l_sql & " LEFT  JOIN evatipoobjpor ON evatipoobj.evatipobjnro = evatipoobjpor.evatipobjnro"
l_sql = l_sql & "		 AND evatipoobjpor.evaevenro = " & l_evaevenro
l_sql = l_sql & " LEFT  JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
l_sql = l_sql & " WHERE evaluaobj.evldrnro	  = " & l_evldrnro
l_sql = l_sql & "   AND evaluaobj.evaborrador = 0 "
l_sql = l_sql & " ORDER BY evatipoobj.evatipobjorden "
rsOpen l_rs, cn, l_sql, 0 
l_evatipobjnro=""
l_cantidad = 0
if l_rs.eof then
%>
<tr>
	<td colspan="<%=l_col%>"><b>No hay Compromisos Cargados.-</b></td>
</tr>
<%
end if
do until l_rs.eof

	if l_evatipobjnro <> l_rs("evatipobjnro") then
		l_evatipobjnro= l_rs("evatipobjnro") 
		l_cantidad = 1 %>
		<tr>
			<td colspan="<%=l_col%>"><b><%=l_rs("evatipobjdabr")%>&nbsp;<%=l_rs("evatipopor")%>%</b></td>
		</tr>
	<%else
		l_cantidad = l_cantidad + 1
	end if%>
    <tr>
        <td align=center valign=top style="width:10"><b><%=l_cantidad%></b></td>
		<td width="200" align=center valign=top><b><%=l_rs("evaobjdext")%></b></td>
		<td align=center>
			<input type="hidden" name="evatipobjnro<%=l_rs("evaobjnro")%>" value="<%=l_rs("evatipobjnro")%>">
			<textarea style="background : #e0e0de;" readonly class="blanc" name="evaobjformed<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=30 rows=5><%=trim(l_rs("evaobjformed"))%></textarea>
		</td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" pond="<%=l_rs("evaobjnro")%>" type="text" name="evaobjpond<%=l_rs("evaobjnro")%>" size=4 value="<%=l_rs("evaobjpond")%>" >
		</td>
        <td align=center>
			<input class="rev" style="background : #e0e0de;" readonly type="text" name="evaobjalcanza<%=l_rs("evaobjnro")%>" size=3 maxlength=3 value="<%=PasarComaAPunto(l_rs("evaobjalcanza"))%>">
			&nbsp;<a href=# onclick="alert(resultados);">?</a>
		</td>
        <td align=center>
			<input readonly class="rev" style="background : #e0e0de;" type="text" name="puntacion<%=l_rs("evaobjnro")%>" size=4 >
			&nbsp;&nbsp;
		</td>
        <td align=center>
			<textarea class="rev" style="background : #e0e0de;" readonly name="evasgtotexto<%=l_rs("evaobjnro")%>"  maxlength=200 size=200 cols=20 rows=4><%=trim(l_rs("evasgtotexto"))%></textarea>
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
        <td align=right><b>Total:</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion<%=l_evatipobjnro%>" size=5>
		</td>
		<td align=right><b>Nota Final:</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="total" type="text" name="nota<%=l_evatipobjnro%>" size=5>
			&nbsp;<a href=# onclick="alert(notafinal);">?</a>
		</td>
		<td><div id="correspondencia<%=l_evatipobjnro%>"></div></td>
		<td>&nbsp;</td>
		</tr>
		<%end if
	else
		%>
		<!-- t o t a  l e s ----------------------------------->
		<tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td align=right><b>Total:</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="blanc" type="text" name="totalponderacion<%=l_evatipobjnro%>" size=4>
		</td>
		<td align=right><b>Nota Final:</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="total" type="text" name="nota<%=l_evatipobjnro%>" size=4>
			&nbsp;<a href=# onclick="alert(notafinal);">?</a>
		</td>
		<td><div id="correspondencia<%=l_evatipobjnro%>"></div></td>
		<td>&nbsp;</td>
		</tr>
		<%
	end if
loop
l_rs.Close
set l_rs = Nothing
%>
	<!-- t o t a  l e s ----------------------------------->
    <tr>
		<td colspan=8>&nbsp;<br></td>
    </tr>
    <tr>
    	<td colspan=5 align=right><b><font size=3>Nota Final Ponderada:</b></td>
		<td align=center>
			<input style="background : #e0e0de;" readonly class="total" type="text" name="notafinalponderada" size=4>
			&nbsp;&nbsp;
		</td>
		<td><div id="correspondencianotafinal"></div></td>
		<td>&nbsp;</td>
    </tr>
<%end if '  SI SE PODIA MOSTRAR

cn.Close
set cn = Nothing
%>
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->


</form>
</body>
</html>
