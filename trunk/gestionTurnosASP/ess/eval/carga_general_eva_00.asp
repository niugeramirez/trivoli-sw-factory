<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=====================================================================================
'Archivo  : carga_general_eva_00.asp
'Objetivo : ABM de objetivos de evaluacion
'Fecha	  : 17-05-2004
'Autor	  : CCRossi
'Fecha	  : 28-07-2004 - Cambiar el tema del puuntaje y opnerlo en evacab. Uno solo.
'Modificación : CCRossi - 02-11-2004 Agregar tablita de referencia abajo
'Modificacion: 19-11-2004-CCRossi-  control de caraceteres raros
'              24-11-2004 CCRossi - Control de caracteres raros
'              13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			   24/05/2007 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'=====================================================================================
 
' Variables
' de uso local  
  Dim l_existe  
  Dim l_visfecha
  Dim l_visdesc
  dim l_evacabnro
  dim l_lista  
  dim l_caracteristica  
  dim l_nombre
  dim l_puntaje
  dim l_puntajemanual
  dim l_primero      

 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 

  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evldrnro
  Dim l_evaseccnro
  dim l_empleg
  
' parametros de entrada---------------------------------------  
  l_evldrnro   = Request.QueryString("evldrnro")
  l_evaseccnro = Request.QueryString("evaseccnro")

 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	


'response.write("<script>alert('"& l_empleg&"')</script>")

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


'__________________________________________________________________________

'							 B O D Y
'__________________________________________________________________________


'buscar la evacab
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evacabnro,empleado.empleg  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
	l_evacabnro = l_rs1("evacabnro")
	l_evaluador = l_rs1("empleg")
 end if
 l_rs1.close
 set l_rs1=nothing

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT puntaje, puntajemanual "
l_sql = l_sql & " FROM  evacab "
l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_puntaje= l_rs("puntaje")
	l_puntajemanual= l_rs("puntajemanual")
end if
l_rs.close
set l_rs=nothing
	
if trim(l_puntajemanual)="" or isnull(l_puntajemanual) then
	l_puntajemanual= l_puntaje
else
	if cdbl(l_puntajemanual)=0 then
		l_puntajemanual= l_puntaje
	end if	
end if
'Crear registros de evaVistos para todos los evldrnro ---------------------------

 Set l_rs = Server.CreateObject("ADODB.RecordSet")	
 l_sql = "SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom, habilitado, "
 l_sql = l_sql & " evatipevalua.evatevdesabr, evadetevldor.evldrnro, "
 l_sql = l_sql & " evaoblieva.evaobliorden, evacab.cabaprobada "
 l_sql = l_sql & " FROM evacab "
 l_sql = l_sql & " inner join evadetevldor on evacab.evacabnro= evadetevldor.evacabnro "
 l_sql = l_sql & " inner join evaoblieva on evaoblieva.evatevnro= evadetevldor.evatevnro "
 l_sql = l_sql & " left join empleado on evadetevldor.evaluador= empleado.ternro "
 l_sql = l_sql & " inner join evatipevalua on evadetevldor.evatevnro= evatipevalua.evatevnro "
 l_sql = l_sql & " inner join evasecc on evadetevldor.evaseccnro= evasecc.evaseccnro "
 l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro
 l_sql = l_sql & " AND evadetevldor.evaseccnro=" & l_evaseccnro
 l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden " 
'response.write l_sql & "<br>"
 rsOpen l_rs, cn, l_sql, 0 
 l_lista="0"
 do until l_rs.eof
   l_lista = l_lista & "," & l_rs("evldrnro")
   
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
	l_sql = l_sql & " FROM  evavistos "
	l_sql = l_sql & " WHERE evavistos.evldrnro   = " & l_rs("evldrnro")
	rsOpen l_rs1, cn, l_sql, 0
	'response.write(l_sql)
	if l_rs1.EOF then
		l_visfecha = cambiafecha(Date(),"","")
		set l_cm = Server.CreateObject("ADODB.Command")
		l_sql = "insert into evavistos "
		l_sql = l_sql & "(evldrnro,visfecha) "
		l_sql = l_sql & "values (" & l_rs("evldrnro") &","&l_visfecha &")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs1.Close
	set l_rs1=nothing
	
   l_rs.MoveNext
 loop 
 l_rs.Close
 set l_rs=nothing 

%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaci&oacute;n General y Comentarios - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
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

function aumentar(anterior, nuevo)
{
	
	if (Number(nuevo.value)< Number(anterior.value))
		nuevo.value=Number(nuevo.value) + 0.5;
	else	
	if (Number(anterior.value)< Number(5))
		nuevo.value=Number(anterior.value) + 0.5;
	else
	{
		alert('La Puntuación no puede superar a 5.')
		nuevo.focus();
	}
	
}
function disminuir(anterior, nuevo)
{
	if (Number(nuevo.value)>Number(anterior.value))
		nuevo.value=Number(nuevo.value) - 0.5;
	else	
	if (Number(anterior.value)> Number(0))
		nuevo.value=Number(anterior.value) - 0.5;
	else
	{
		alert('La Puntuación no puede ser inferior a 0.')
		nuevo.focus();
	}
	
}
function validar(fecha, texto)
{
	if (!(validarfecha(fecha))) 
	{
		error=true;
		fecha.focus();
		return false;
	}	
	else	
	texto.value= Blanquear(texto.value)
	if (texto.value.length>200){
		alert('La Observación no puede superar 200 caracteres.')
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
<style>
.rev
{
	font-size: 10;
	border-style: none;
}
.blanc
{
	font-size: 10;
	border-style: none;
	background : transparent;
}
</style>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos">

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<td colspan="4" align="left" class="th2">Evaluaci&oacute;n General y Comentarios</td>
<tr>
	
<%	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evavistos.evldrnro, visdesc,visfecha, evatevdesabr , empleado.empleg"
	l_sql = l_sql & " FROM  evavistos "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evavistos.evldrnro"
	l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro = evadetevldor.evatevnro"
	l_sql = l_sql & " WHERE evavistos.evldrnro IN (" & l_lista & ")"
	rsOpen l_rs1, cn, l_sql, 0
	l_primero=-1
	do while not l_rs1.eof
	    l_evaluador = l_rs1("empleg")
	    
		if Int(l_evldrnro) <> l_rs1("evldrnro") then
			l_caracteristica = "readonly disabled"
			l_nombre = l_evldrnro
		else	
			l_caracteristica = ""
			l_nombre = ""
		end if
		if l_primero=-1 then%>
		<tr>
		<td colspan=4>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align=right valign=center><b>Puntuaci&oacute;n Obtenida</b></td>
				<td valign=center><input <%=l_caracteristica%> readonly style="background : #e0e0de;" class="rev" type="text" name="puntaje" size="2" maxlength="2" value="<%=PasarComaAPunto(l_puntaje)%>"></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td align=right valign=top><b>Evaluaci&oacute;n General</b></td>
				<td valign=top>
					<input readonly style="background : #e0e0de;" class="rev" type="text" name="puntajemanual" size="2" maxlength="2" value="<%=PasarComaAPunto(l_puntajemanual)%>">
					&nbsp;
					<a href="Javascript:aumentar(document.datos.puntaje,document.datos.puntajemanual);"><font type="tahoma" size=1>Aumentar</font></a>&nbsp;
					<a href="Javascript:disminuir(document.datos.puntaje,document.datos.puntajemanual);"><font type="tahoma" size=1>Bajar</font></a>
				</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			</table>	
		</td>
		</tr>
	<%
	l_primero=0
	end if
	
	%>
	<tr>
		<td width="10%">
			<b><%=l_rs1("evatevdesabr")%></b>
		</td>
		<td width="50%">
			<textarea <%=l_caracteristica%> name="visdesc<%=l_nombre%>"  maxlength=200 size=200 cols=50 rows=4><%=trim(l_rs1("visdesc"))%></textarea>
		</td>
		<td width="25%">
			<input <%=l_caracteristica%> type="text" name="visfecha<%=l_nombre%>" size="10" maxlength="10" value="<%=l_rs1("visfecha")%>">
			<%if trim(l_caracteristica)="" then %>
			<a href="Javascript:Ayuda_Fecha(document.datos.visfecha<%=l_nombre%>)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
			<%end if%>
		</td>
		<td valign=middle width="15%">
			<%if trim(l_caracteristica)="" then %>
			<a href=# onclick="if (validar(document.datos.visfecha<%=l_nombre%>,document.datos.visdesc<%=l_nombre%>) ) {grabar.location='grabar_vistos_evaluacion_00.asp?tipo=M&evldrnro=<%=l_evldrnro%>&visdesc='+escape(Blanquear(document.datos.visdesc.value))+'&visfecha='+document.datos.visfecha.value+'&puntaje='+document.datos.puntajemanual.value;document.datos.grabado.value='M';}">Grabar</a>			
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado" size="1">
			<%end if%>
		</td>
    </tr>
    <%l_rs1.MoveNext
   loop
   l_rs1.close
   set l_rs1=nothing
%>
</form>	
	<tr>
		<td colspan=4>&nbsp;
		</td>
    </tr>
	<tr>
		<td width="30%">&nbsp;</td>
		<td colspan=2 align=center colspan=4>
			<table border=1 >
			<tr>
				<td><b>Puntuaci&oacute;n</b></td>
				<td><b>Correspondencia Literal</b></td>
			</tr>
			<tr>
				<td>0&nbsp;1</td>
				<td>Significativamente debajo de los requerimientos</td>
			</tr>
				<td>1,5&nbsp;;&nbsp;2&nbsp;y&nbsp;2,5</td>
				<td>Debajo de los requerimientos</td>
			<tr>
				<td>3</b></td>
				<td>Cumple con los requerimientos</td>
			</tr>
			<tr>
				<td>3,5&nbsp;y&nbsp;4</b></td>
				<td>Encima de los requerimientos</td>
			</tr>
			<tr>
				<td>4,5&nbsp;y&nbsp;5</b></td>
				<td>Supera significativamente los requerimientos</td>
			</tr>
			</table>
		</td>
		<td width="20%">&nbsp;</td>
    </tr>
</table>
<iframe  src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
<!--iframe name="grabar" style="width:500;height:100"-->
</iframe>

</body>
</html>
