<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: ver_resultadosyarea_eva_00.asp
'Descripción	: ver competencias y resultado total por area
'Autor			: 27-09-2004
'Fecha			: CCRossi
'Modificado		: 08-03-2005 LAmadio - cambiar el fuente para que liste todas las areas y su respectivas competencias.
'            	  13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================

on error goto 0
' Variables
 
' de uso local  
  Dim l_evafacnro
  Dim l_evatrnro
  Dim l_evatitdesabr
  Dim l_observables
  Dim l_interpretaciones

  dim l_areatrnro
  dim l_evaareadesc

  dim l_areaResu
  dim l_areaDescrip
  dim l_compResu
  
  dim l_compDescrip
  dim l_estrnro
  dim l_ternro  
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

 dim l_evatevnro
 dim l_sinarea
  
' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")

  ' _____________________________________________________________________
sub datosArea (areaResu, areaDescrip)
areaResu=""
areaDescrip=""

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evatrnro, evaareadesc  "
l_sql = l_sql & " FROM evaarea "
l_sql = l_sql & " WHERE evaarea.evatitnro  = " & l_rs("evatitnro")
l_sql = l_sql & " AND   evaarea.evldrnro  = " & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0
if not l_rs1.eof then
	l_areatrnro=l_rs1("evatrnro")
	l_evaareadesc=l_rs1("evaareadesc")
end if
l_rs1.close
set l_rs1=nothing
			
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaresu.evatrnro,  "
l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaresu.evafacnro "
l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
l_sql = l_sql & " AND   evatitulo.evatitnro  = " & l_rs("evatitnro")
l_sql = l_sql & " order by evatrvalor "
rsOpen l_rs1, cn, l_sql, 0

'if l_sinarea="SI" then
	areaResu= areaResu & "<select name=areatrnro"& l_rs("evatitnro")&" disabled>"
'else 
	'areaResu= areaResu & "<select name=areatrnro"& l_rs("evatitnro")&">"
'end if
areaResu= areaResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"

do while not l_rs1.eof
	areaResu= areaResu & "<option value="&l_rs1("evatrnro")&">"&l_rs1("evatrvalor")& "&nbsp;&nbsp;&nbsp;"& l_rs1("evatrdesabr")&"</option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
areaResu= areaResu & "	</select> "

areaResu= areaResu & " <script>document.datos.areatrnro"&l_rs("evatitnro")&".value='"&l_areatrnro&"';</script>"

'if l_sinarea="SI" then
	areaDescrip = "<textarea name=evaareadesc"&l_rs("evatitnro")&" cols=30 rows=4 disabled>" & trim(l_evaareadesc)& "</textarea>"
'else
	'areaDescrip = "<textarea name=evaareadesc"&l_rs("evatitnro")&" cols=30 rows=4>" & trim(l_evaareadesc) & "</textarea>"
'end if
end sub   

' ______________________________________________________________
sub datosCompetenc (compResu,interpr, compDescrip) 'buscar evaresu
compResu = ""
compDescrip = ""

Set l_rs1 = Server.CreateObject("ADODB.RecordSet") 
l_sql = "SELECT evaresu.evatrnro, evaresu.evaresudes, "
l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
l_sql = l_sql & " FROM evaresu "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rs("evafacnro")
l_sql = l_sql & " order by evatrvalor "
rsOpen l_rs1, cn, l_sql, 0

compResu = compResu & "<select name=evatrnro"& l_rs("evafacnro")&" disabled>"
compResu = compResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
interpr=""
do while not l_rs1.eof
	if trim(l_rs1("evaresudes")) <>"" and not isnull(l_rs1("evaresudes")) then
		interpr = l_interpretaciones & l_rs1("evatrdesabr") &": "&l_rs1("evaresudes") & "\n"
	end if
	compResu = compResu & "<option value="& l_rs1("evatrnro")&">"& l_rs1("evatrvalor")&"&nbsp;&nbsp;&nbsp;"&l_rs1("evatrdesabr")&" </option>"
l_rs1.MoveNext
loop 
l_rs1.Close
set l_rs1 = nothing
compResu = compResu & "</select>"

compResu = compResu & "<script>document.datos.evatrnro"&l_rs("evafacnro")&".value='"& l_rs("evatrnro")&"'</script>"

compDescrip = "	<input type=hidden name=evaresudesc"&l_rs("evafacnro")&">"
'compDescrip = "	<textarea name=evaresudesc"&l_rs("evafacnro")&" cols=30 rows=5 disabled>"& trim(l_rs("evaresudesc"))&"</textarea>"

 end sub

 '__________________________________________________
sub	datosObservables(obs)
obs = ""
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evadescomp.evadcdes, estructura.estrdabr "
l_sql = l_sql & " FROM evadescomp "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evadescomp.estrnro "
l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs("evafacnro")
'l_sql = l_sql & " AND   evadescomp.tenro     =  " & ctenro
l_sql = l_sql & " AND   evadescomp.estrnro IN (" & l_estrnro & ")"
rsOpen l_rs1, cn, l_sql, 0

l_observables=""
do while not l_rs1.eof
	obs = obs & l_rs1("estrdabr") & " - "& l_rs1("evadcdes")& "\n"
	l_rs1.MoveNext
loop
l_rs1.Close
set l_rs1 = nothing
end sub

'__________________________________________________________________________________ 
 '  verrrrrrrrrrrrrrrrrrrrrrrrr

'Buscar el rol de evaluador que esta entrando
    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evatevnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
   l_evatevnro= l_rs1("evatevnro")
 end if 
 l_rs1.close
 set l_rs1=nothing
 l_sinarea="NO"
 if (l_evatevnro <> cevaluador) then ' 6 o ?????? - vERRR si pregunta por logueado cdo es revisor y auto
	l_sinarea="SI"
 end if
 
 
' Crear registros de evaresultado para los facnro y el evldrnro
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evaseccfactor.evaseccnro, evaseccfactor.evafacnro, evaresu.evatrnro, "
  l_sql = l_sql & " evatitulo.evatitnro "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " INNER JOIN evatitulo ON evatitulo.evatitnro = evafactor.evatitnro "
  l_sql = l_sql & " INNER JOIN evaresu   ON evaresu.evaseccnro  = evaseccfactor.evaseccnro AND  evaresu.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  rsOpen l_rs, cn, l_sql, 0
  'response.write l_sql
  set l_cm = Server.CreateObject("ADODB.Command")  
  
  if not l_rs.eof then
		l_evatrnro  = "null"
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT *  "
		l_sql = l_sql & " FROM  evaarea "
		l_sql = l_sql & " WHERE evaarea.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " AND   evaarea.evatitnro  = " & l_rs("evatitnro")
		rsOpen l_rs1, cn, l_sql, 0
		if l_rs1.EOF then
			l_sql = "INSERT INTO evaarea "
			l_sql = l_sql & " (evldrnro, evatitnro, evatrnro, evaareadesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_rs("evatitnro")	 & ","
			l_sql = l_sql & l_evatrnro & ",'')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs1.Close
		set l_rs1=nothing
  end if
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
		l_rs1.Close
		set l_rs1=nothing
		l_rs.MoveNext
		
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
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Competencias de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
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
</script>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >
<form name="datos">
<input type="Hidden" name="terminarsecc" value="SI">

<table border="0" cellpadding="0" cellspacing="1" height="100%" width="100%">
<%'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   Set l_rs = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaresultado.evldrnro, evaresultado.evafacnro, "
   l_sql = l_sql & " evaresultado.evatrnro, evaresultado.evaresudesc, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
   l_sql = l_sql & " evatitulo.evatitnro, evatitulo.evatitdesabr, "
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
   		<tr style="height:5">
			<th align=left colspan="5" class="th2"><b><%=l_rs("evatitdesabr")%></b></th></tr>
		<tr>
			<td align="right"><b>Resultado de Area</b>&nbsp;</td>	
			<% datosArea l_areaResu, l_areaDescrip 'buscar evaarea %>
			<td align="left" valign="middle"> <%=l_areaResu%></td>
			<td align="left"><%=l_areaDescrip %></td>
			<td colspan=2 nowrap>
				<%'if l_sinarea<>"SI" then%>
				<!--
				<a href=# onclick="if (Controlar(document.datos.areatrnro<%=l_rs("evatitnro")%>)) {grabar.location='grabar_areas_evaluacion_00.asp?evatitnro=<%=l_rs("evatitnro")%>&evldrnro=<%=l_evldrnro%>&evaareadesc='+escape(Blanquear(document.datos.evaareadesc<%=l_rs("evatitnro")%>.value))+'&evatrnro='+document.datos.areatrnro<%=l_rs("evatitnro")%>.value;document.datos.grabado<%=l_rs("evatitnro")%>.value='G'; }">Grabar</a>
				<br>
				<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evatitnro")%>" size="1">
				-->
				<%'end if%>
				&nbsp;
			</td>
		</tr>
		<tr>	<td colspan="5">&nbsp;</td>		</tr>
		<tr>
			<td colspan="2"><b>Competencias </b></td>
			<td><b>Resultado</b></td>
			<!-- <td> <b>Observaci&oacute;n</b></td> -->
			<td>&nbsp;</td>
			<td><b>Observables</b></td>
		</tr>
<%		l_evatitdesabr = l_rs("evatitdesabr")
	end if %>
		<tr>
		<% datosCompetenc l_compResu, l_interpretaciones,l_compDescrip   'buscar evaresu 	%>
		<td valign="top" colspan="2"><%=l_rs("evafacdesabr")%></td>	
		<td nowrap valign="top"> 
			<%=l_compResu%>
			<%if trim(l_interpretaciones)="" then%>
				<a href=# onclick="alert('No hay Interpretaciones cargadas para estos resultados.')">?</a></td>
			<%else%>	
				<a href=# onclick="alert('<%=l_interpretaciones%>')">?</a></td> 
			<%end if%>	
			<%=l_compDescrip%>
		</td>
		<!-- <td valign="top">	</td> -->
		<td nowrap valign="top">
			&nbsp;
			<!--
			<a href=# onclick="if (Controlar(document.datos.evatrnro<%=l_rs("evafacnro")%>)){ grabar.location='grabar_competencias_evaluacion_00.asp?evafacnro=<%=l_rs("evafacnro")%>&evldrnro=<%=l_evldrnro%>&evaresudesc='+escape(Blanquear(document.datos.evaresudesc<%=l_rs("evafacnro")%>.value))+'&evatrnro='+document.datos.evatrnro<%=l_rs("evafacnro")%>.value;document.datos.grabado<%=l_rs("evafacnro")%>.value='G'; }">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evafacnro")%>" size="1">
			-->
		</td>
<% 		datosObservables l_observables
		if trim(l_observables)="" then%>
			<td valign=top align=center><a href=# onclick="alert('No hay definidas Conductas Observables \n para las Estructuras del Empleado \n y la Competencia.')">?</a></td>
		<%else%>	
			<td valign=top align=center><a href=# onclick="alert('<%=l_observables%>')">?</a></td>
		<%end if%>	
	</tr>
<%
	l_rs.Movenext
  loop
l_rs.Close  %>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
