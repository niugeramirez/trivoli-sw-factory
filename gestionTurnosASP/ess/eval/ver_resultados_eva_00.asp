<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
' Objetivo  : VISUALIZACION DE RESULTADOS COMPETENCIA-READ ONLY
' Modificado: 10-05-2004  CCRossi agregar titulo de evatitulo cuando cambia de area
' Modificado: 08-09-2004  CCRossi Cambiar consulta de Conductas Observables. 
'					      Que muestren todas las qu etengan por estructura.
' Modificado: 08-09-2004  CCRossi Agregar consulta de interpretacion de todos 
'					      los resultados posibles para la competencia.
' Modificado: 04-10-2004 CCRossi Agregar campo evaresuEJEM, para ABN...
' Modificado: 22-10-2004 CCRossi poner "(mi borrador)" si cejemp=-1 ( o sea es ABN....)
' Modificado: 22-10-2004 CCRossi no mostrar Observables si cejemp=-1 ( o sea es ABN....)
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 18-08-2006 - LA. Ordenar por Areas y agregar funcion unescape()
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'============================================================================================

' Variables
 
' de uso local  
  Dim l_evafacnro
  Dim l_evatrnro
  Dim l_evatitdesabr
  Dim l_observables
  Dim l_interpretaciones
    
  dim l_estrnro
  dim l_ternro  
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

' Crear registros de evaresultado para los facnro y el evldrnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evaseccfactor.evaseccnro, evaseccfactor.evafacnro, evaresu.evatrnro "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
  l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0
  'response.write l_sql
  set l_cm = Server.CreateObject("ADODB.Command")  
  
  do while not l_rs.eof
		l_evafacnro = l_rs("evafacnro")
		l_evatrnro  = "null"
  
  		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & " FROM  evaresultado "
		l_sql = l_sql & " WHERE evaresultado.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " AND   evaresultado.evafacnro  = " & l_rs("evafacnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evaresultado "
			l_sql = l_sql & " (evldrnro, evafacnro, evatrnro, evaresudesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_rs("evafacnro")	 & ","
			l_sql = l_sql & l_evatrnro & ",'')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
		end if
		l_rs.MoveNext
		l_rs1.Close
  loop
  l_rs.Close
  set l_rs=nothing
  
'buscar el ternro del EVALUADO --------------------------------------------------------
l_ternro=""
l_estrnro="0"
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado  "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " WHERE evadetevldor.evldrnro =" & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_ternro = l_rs("empleado")
end if	
l_rs.Close
set l_rs=nothing


'buscar las estructuras ACTIVAS del empleado -----------------------------------------------------------------
if trim(l_ternro) <> "" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT his_estructura.estrnro, teorden "
	l_sql = l_sql & " FROM his_estructura "
	l_sql = l_sql & " INNER JOIN estructura		ON estructura.estrnro=his_estructura.estrnro "
	l_sql = l_sql & " INNER JOIN tipoestructura ON tipoestructura.tenro=estructura.tenro "
	l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
	l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
	l_sql = l_sql & " ORDER BY teorden " 
	rsOpen l_rs, cn, l_sql, 0
	do while not l_rs.eof 
		l_estrnro = l_estrnro & "," & l_rs("estrnro")
		l_rs.MoveNext
	loop	
	l_rs.Close
	set l_rs=nothing
end if

' MOSTRAR evaresudes dependiendo del valor que elija como resultado -----
response.write "<script languaje='javascript'>" & vbCrLf
response.write "function Mostrar(evatrnro,evafacnro){ " & vbCrLf
response.write "};" & vbCrLf
response.write "</script>" & vbCrLf
'Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
'l_sql = "SELECT evaresu.evatrnro, evaresu.evafacnro, evaresu.evaresudes "
'l_sql = l_sql & " FROM evaresu "
'l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
'rsOpen l_rs1, cn, l_sql, 0 
'dim i 
'i = 0
'do while not l_rs1.eof
'		response.write "if ((evatrnro == " & l_rs1(0) & ") && (evafacnro == " & l_rs1(1) & ") ) {" & vbCrLf
'		response.write "document.datos.evaresudes" & l_rs1("evafacnro")& ".value = '" & l_rs1(2) & "';" & vbCrLf
'		response.write "return '" & l_rs1(2) & "';" & vbCrLf
'		response.write "};" & vbCrLf
'l_rs1.MoveNext
'loop
'l_rs1.Close
'set l_rs1 = nothing
'CONTROL DE EVALUAOR LOGEADO =================================================================
 dim l_empleg
 dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
 dim l_mostrar '1 o 0 si tiene que mostrar la observacion. 
 l_empleg = Session("empleg")
 if trim(l_empleg)="" then
	l_empleg = Request.QueryString("empleg")
 end if	
 
'buscar la evacab
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT v_empleado.empleg  "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN v_empleado ON v_empleado.ternro = evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.EOF then
	l_evaluador = l_rs("empleg")
 end if
 l_rs.close
 set l_rs=nothing
 
'Response.Write l_empleg & "<br>" & l_evaluador
 l_mostrar = "0"
 if trim(l_empleg)<>"" and not isNull(l_empleg) then
	if trim(l_empleg) = trim(l_evaluador) then
		l_mostrar = "1"
	else	
		l_mostrar = "0"
	end if
 else
 	l_mostrar = "1"
 end if
' ==============================================================================================
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Competencias de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<style>
.rev
{
	font-size: 10;
	border-style: none;
}
</style>
</head>

<script>
function Controlar(resu){
	if ((resu.value=="")||(resu.value=="0")){
		alert('Seleccione un resultado.');
		resu.focus();
		return false;
	}
	else
		return true;
}	

function Promedio(){
	var r = showModalDialog('calcular_promediocomp_eva_00.asp?evldrnro=<%=l_evldrnro%>&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:5;dialogHeight:5'); 
	document.datos.promedio.value=r;
}

</script>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Promedio();">
<form name="datos">

<table border="0" cellpadding="1" cellspacing="1" height="100%" width="100%">
<!--tr style="border-color :CadetBlue;">
	<td colspan="5" align="left" class="barra">Carga de Competencias de Evaluaci&oacute;n</td>
<tr-->

<%'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaresultado.evldrnro, evaresultado.evafacnro,  evatitdesabr ,"
   l_sql = l_sql & " evaresultado.evatrnro, evaresultado.evaresudesc, evaresultado.evaresuejem, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
   l_sql = l_sql & " evatitulo.evatitdesabr, "
   l_sql = l_sql & " evaseccfactor.orden "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
   l_sql = l_sql & " ORDER BY evatitulo.evatitdesabr, evaseccfactor.orden "
   l_evatitdesabr=""
   rsOpen l_rs, cn, l_sql, 0
   do while not l_rs.eof 
   if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then%>
	<tr style="border-color :CadetBlue;" style="height:5">
		<th colspan="6" align="left" class="th2"><%=l_rs("evatitdesabr")%></th>
	<tr>
	<tr style="border-color :CadetBlue;"  style="height:5">
		<td><b>Descripci&oacute;n</b></td>
		<td><b>Puntuaci&oacute;n</b></td>
		<%if cejemplo=-1 then%>
		<td><b>Comentario<br>
		Explicación <br>
		Ejemplos de comportamientos observados</b></td>
		<%end if%>
		<td align=center><b>Observaciones <%if cejemplo=-1 then%>(mi borrador)<%end if%></b></th>
		<%if cejemplo<>-1 then%>
		<td><b>Observables</b></td>
		<%end if%>
	</tr>
	<%	l_evatitdesabr = l_rs("evatitdesabr")
		end if%>
	<tr>
		<td valign="top"><%if trim(l_rs("evafacdesext"))="" or isnull(l_rs("evafacdesext")) then%> <%=l_rs("evafacdesabr")%> 
						<%else%>
						<%=l_rs("evafacdesext")%> 
						<%end if%>
						</td>	
		<td nowrap valign="top">
			<%'BUSCAR la descripcion de evaresu  ----------------------------
		    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaresu.evatrnro, evaresu.evaresudes, "
			l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
			l_sql = l_sql & " FROM evaresu "
			l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
			l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
			l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rs("evafacnro")
			l_sql = l_sql & " order by evatrvalor "
			rsOpen l_rs1, cn, l_sql, 0
			%>
			<select readonly disabled name="evatrnro<%=l_rs("evafacnro")%>">
			<script>
			// onchange="Mostrar(document.datos.evatrnro<%'=l_rs("evafacnro")%>.value,<%'=l_rs("evafacnro")%>);"
			</script>
			<option value=0> 0&nbsp;&nbsp; Sin Evaluar</option>
			<%l_interpretaciones=""
			  do while not l_rs1.eof
				if trim(l_rs1("evaresudes")) <>"" and not isnull(l_rs1("evaresudes")) then
					l_interpretaciones = l_interpretaciones & l_rs1("evatrdesabr") &": "&l_rs1("evaresudes") & "\n"
				end if%>
				<option value=<%=l_rs1("evatrnro")%>><%=l_rs1("evatrvalor")%>&nbsp;&nbsp;&nbsp;<%=l_rs1("evatrdesabr")%></option>
			<%l_rs1.MoveNext
			loop 
			l_rs1.Close
			set l_rs1 = nothing%>
			</select>
			<!--input  disabled type="text" name="evaresudes<%'=l_rs("evafacnro")%>"-->
			<script>document.datos.evatrnro<%=l_rs("evafacnro")%>.value='<%=l_rs("evatrnro")%>'</script>
			<!--script>Mostrar(document.datos.evatrnro<%'=l_rs("evafacnro")%>.value,<%'=l_rs("evafacnro")%>);</script-->
			<%if trim(l_interpretaciones)="" then%>
				<a href=# onclick="alert('No hay Interpretaciones cargadas para estos resultados.')">?</a></td>
			<%else%>	
				<a href=# onclick="alert('<%=unescape(l_interpretaciones)%>')">?</a></td>
			<%end if%>	
		</td>
		<%if cejemplo=-1 then%>
		<td valign="top">
			<textarea  readonly disabled  name="evaresuejem<%=l_rs("evafacnro")%>" cols=25 rows=4><%=trim(l_rs("evaresuejem"))%></textarea>
		</td>
		<%else%>
		<input type="hidden" name="evaresuejem<%=l_rs("evafacnro")%>">
		<%end if%>
		<td valign="top">
			<%if l_mostrar="1" then%>
			<textarea readonly disabled name="evaresudesc<%=l_rs("evafacnro")%>" cols=25 rows=4><%=trim(l_rs("evaresudesc"))%></textarea>
			<%else%>
			<input type="hidden" name="evaresudesc<%=l_rs("evafacnro")%>" size=5 value="<%=trim(l_rs("evaresudesc"))%>">
			No habilitado.
			<%end if%>
		</td>
		
		</td>
		
		<%  Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evadescomp.evadcdes, estructura.estrdabr "
			l_sql = l_sql & " FROM evadescomp "
			l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evadescomp.estrnro "
			l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs("evafacnro")
			'l_sql = l_sql & " AND   evadescomp.tenro     =  " & ctenro
			l_sql = l_sql & " AND   evadescomp.estrnro IN (" & l_estrnro & ")"
			rsOpen l_rs1, cn, l_sql, 0
			l_observables=""
			do while not l_rs1.eof
				l_observables = l_observables & l_rs1("estrdabr") & " - "& l_rs1("evadcdes")& "\n"
				l_rs1.MoveNext
			loop
			l_rs1.Close
			set l_rs1 = nothing
			
			if cejemplo<>-1 then
			if trim(l_observables)="" then%>
				<td valign=top align=center><a href=# onclick="alert('No hay definidas Conductas Observables \n para las Estructuras del Empleado \n y la Competencia.')">?</a></td>
			<%else%>	
				<td valign=top align=center><a href=# onclick="alert('<%=unescape(l_observables)%>')">?</a></td>
			<%end if
			end if%>	
		</tr>
		<%
		
		l_rs.Movenext
		loop
		l_rs.Close%>
	<!-- Promedio ----------------------------------->
    <tr style="height:10">
		<td align=right><b>Promedio</b></td>
		<td align=left>
		<input style="background:#e0e0de;" readonly class="blanc" type="text" name="promedio" size=5></td>
		<td align=center colspan="<%if cejemplo<>-1 then%>4<%else%>3<%end if%>"></td>
		
    </tr>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
