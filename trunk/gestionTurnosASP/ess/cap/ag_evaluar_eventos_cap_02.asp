<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
Archivo: 		ag_evaluar-eventos_cap_02.asp
Descripción: 	Muestra el encabezado de la evaluacion de Empleados a los Eventos
Autor : 		Raul Chinestra
Fecha: 			25/06/2007
-->
<% 
on error goto 0

Dim l_rs
Dim l_cm
Dim l_sql
Dim l_ternro
Dim l_tesnro
Dim l_pasnro
Dim l_codigo
Dim l_testdesc
Dim l_tercero
Dim l_fornro
Dim l_fordesabr

Dim l_ttesnro
Dim l_evenro
Dim l_ttesfor

Dim l_tesfec
Dim l_teshor
Dim l_tesfin
Dim l_testie
Dim l_primeravez
l_primeravez = true

l_ttesfor = 0
l_ttesnro	= request.QueryString("ttesnro")
l_ternro	= request.QueryString("ternro")
l_evenro	= request.QueryString("evenro")

l_ternro = l_ess_ternro

cn.BeginTrans

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")

' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado()
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id","cap_evento")
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()


' ------------------------------------------------------------------------------------------------------------------
' Busco si ya ingreso a realizar el test previamente
' ------------------------------------------------------------------------------------------------------------------
l_sql = "SELECT tesnro, fornro, tesfec, teshor, tesfin, testie " 
l_sql = l_sql & " FROM test "
l_sql = l_sql & " WHERE test.ttesnro = " & l_ttesnro 
l_sql = l_sql & "   AND test.ternro = " & l_ternro
l_sql = l_sql & "   AND test.evenro = " & l_evenro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	'--------------------------------------------------------------------------------------------------------------------
	' Si NO Existe el test
	'--------------------------------------------------------------------------------------------------------------------
	l_rs.Close
	l_sql = "SELECT ttesfor " 
	l_sql = l_sql & " FROM pos_tipotest "
	l_sql = l_sql & " WHERE ttesnro = " & l_ttesnro 
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		if isNull(l_rs("ttesfor")) then
			l_ttesfor = 1
		else
			l_ttesfor = l_rs("ttesfor")
		end if
	else
		l_ttesfor = 1
	end if
	l_rs.close	
	'--------------------------------------------------------------------------------------------------------------------
	' Primero busco el formulario
	'--------------------------------------------------------------------------------------------------------------------
	'	pos_tipotest.ttesfor = 1 random
	'						 = 2 por default
	'--------------------------------------------------------------------------------------------------------------------
		if Cint(l_ttesfor) = Cint(1) then
			'-------------------------------------------------------------------
			'Tomo al azar
			'-------------------------------------------------------------------
			l_sql = "SELECT fornro " 
			l_sql = l_sql & " FROM pos_formulario "
			l_sql = l_sql & " WHERE ttesnro = " & l_ttesnro 
			'rsOpen l_rs, cn, l_sql, 0 
			rsOpencursor l_rs, cn, l_sql, 0, adopenkeyset 
			if not l_rs.eof then
				randomize
				Dim l_nroazar
				l_nroazar = Int((l_rs.RecordCount  * Rnd) + 1)
				l_rs.AbsolutePosition = l_nroazar
				l_fornro = l_rs("fornro")
			end if
		else
			'-------------------------------------------------------------------
			'Tomo el definido por default
			'-------------------------------------------------------------------
			l_sql = "SELECT fornro " 
			l_sql = l_sql & " FROM pos_formulario "
			l_sql = l_sql & " WHERE ttesnro = " & l_ttesnro 
			l_sql = l_sql & " AND fordef = -1 "
			rsOpen l_rs, cn, l_sql, 0
			if not l_rs.eof then
				l_fornro = l_rs("fornro")
			else
				'valido si no hay ninguno por default tomo el primero
				l_rs.close
				l_sql = "SELECT fornro " 
				l_sql = l_sql & " FROM pos_formulario "
				l_sql = l_sql & " WHERE ttesnro = " & l_ttesnro 
				l_sql = l_sql & " AND fordef = -1 "
				rsOpen l_rs, cn, l_sql, 0
				if not l_rs.eof then
					l_fornro = l_rs("fornro")
				end if
			end if
		end if

	l_rs.close	
	'---------------------------------------------------------------------------------------------------------------
	' Creo el Nuevo Test
	'---------------------------------------------------------------------------------------------------------------
	l_sql = "insert into test "
	'l_sql = l_sql & "(ttesnro, ternro, resnro, tesresp , tescosto, tesarchivo, tesobs, entnro,pedbusnro, fornro) "
	l_sql = l_sql & "(ttesnro, ternro, fornro, evenro, tesfec, teshor, tesfin) "
	l_sql = l_sql & "values (" & l_ttesnro 
	l_sql = l_sql & ", " & l_ternro
	'l_sql = l_sql & ", " & l_resnro
	'l_sql = l_sql & ", '" & l_tesresp & "' "
	'l_sql = l_sql & ", " & l_tescosto
	'l_sql = l_sql & ", '" & l_tesarchivo & "' "
	'l_sql = l_sql & ", '" & l_tesobs & "'"
	'l_sql = l_sql & ",  " & l_entnro
	'l_sql = l_sql & ",  " & l_pedbusnro
	l_sql = l_sql & ",  " & l_fornro
	l_sql = l_sql & ",  " & l_evenro
	l_sql = l_sql & ",  " & cambiafecha(date(),"YMD",true)
	l_sql = l_sql & ",  '" & hour(now) & ":" & minute(now) & "'"
	l_sql = l_sql & ",  " & 0
	l_sql = l_sql & ")"	
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	'----------------------------------------------------------------------------------------------------------------
	' Obtengo el codigo interno del test generado
	'----------------------------------------------------------------------------------------------------------------
	l_tesnro = codigogenerado()	
	
	l_tesfec = date
	l_teshor = hour(now) & ":" &  minute(now)
	l_tesfin = 0
	l_primeravez = true
else
	'--------------------------------------------------------------------------------------------------------------------
	' Si Existe el test
	'--------------------------------------------------------------------------------------------------------------------
	l_tesnro = l_rs("tesnro")
	l_fornro = l_rs("fornro")
	l_tesfec = l_rs("tesfec")
	l_teshor = l_rs("teshor")
	if isnull(l_rs("tesfin")) then
		l_tesfin = 0
	else
		l_tesfin = l_rs("tesfin")
	end if
	l_testie = l_rs("testie")
	l_tesfec = l_rs("tesfec")
	l_teshor = l_rs("teshor")
	l_primeravez = false
	l_rs.close
end if

'-------------------------------------------------------------------------------------------------------------------------
' aca se debe buscar el formulario seleccionado del tipo de test
'-------------------------------------------------------------------------------------------------------------------------
l_sql = "SELECT ttesdesabr, fornro, fordesabr " 
l_sql = l_sql & " FROM pos_tipotest "
l_sql = l_sql & " INNER JOIN pos_formulario ON pos_formulario.ttesnro = pos_tipotest.ttesnro "
l_sql = l_sql & " WHERE pos_tipotest.ttesnro = " & l_ttesnro
l_sql = l_sql & " AND fornro = " & l_fornro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_testdesc	= l_rs("ttesdesabr")
	l_fornro    = l_rs("fornro")
	l_fordesabr	= l_rs("fordesabr")	
else
	l_testdesc = "??"
end if
l_rs.close	

'------------------------------------------------------------------------------------------------------------------------
' Busco los datos del Empleados
'------------------------------------------------------------------------------------------------------------------------
l_sql = "SELECT terape, terape2, ternom, ternom2  " 
l_sql = l_sql & " FROM tercero "
l_sql = l_sql & " WHERE ternro = " & l_ternro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_tercero	= l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
end if
l_rs.close
set l_rs = nothing

cn.CommitTrans

%>
<html>
<head>
<link href="../<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<title>Evaluaciones - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Aceptar(){
	<% If l_tesfin = 0 then%>
		finalizar();
	<% Else %>
		window.close();
	<% End If %>
}

function PrevPre(){
	grabar();
	var pre = Number(document.all.pregunta.value);
	if (pre > 1){
		pre = pre - 1;
		document.all.pregunta.value = pre;
		Actualizar();
	}
}

function NextPre(){
	grabar();
	var pre = Number(document.all.pregunta.value);
	var tot = Number(document.all.totpregunta.value);
	if (pre < tot){
		pre = pre + 1;
		document.all.pregunta.value = pre;
		Actualizar();
	}
}

function Actualizar(){
	var param	
	var filtro

	param = "pregunta=" + document.all.pregunta.value + '&tesnro=<%= l_tesnro%>&ternro=<%= l_ternro %>&fornro=<%= l_fornro %>&tesfin=<%= l_tesfin %>';
	document.ifrm.location= 'ag_evaluar_eventos_cap_03.asp?'+param;
} 

function refrescar(){
	document.ifrm.location="vacio.asp";
	//document.all.porpagant.value=0;
	document.all.pagina.value=1;
	document.all.totpagina.value=1;
	document.all.porpagina.value="";
	document.all.totalempl.value=0;
	document.all.porpagina.disabled= true;
}

function grabar(){
	<% If Cint(l_tesfin) <> -1 Then %>
		var allTexts = document.ifrm.document.getElementsByTagName("textarea");
		var i;
		var r;
		var estado;
		estado = "si";
		//valido que ningun textarea supere la maxima longitud permitida.
		for(i=0; i < allTexts.length; i++){
			r = allTexts.item(i).value.length;
			if (r > document.ifrm.MaxTotal){
				alert('Ha llegado al máximo de caracteres.');
				document.ifrm.setSelectionRange(allTexts.item(i), document.ifrm.MaxTotal, allTexts.item(i).value.length)
				estado = "no";
				return;
			}
		}
		// Si todo anda ok.
		if (estado == "si"){
			//document.ifrm.datos.value = escape(document.ifrm.datos.value);
			//abrirVentanaH('','pepe',500,500);
			document.ifrm.datos.target = 'valida';
			document.ifrm.datos.submit();
		}
	<% End If %>
}

//Inicializo la hora
<% If l_primeravez = true Then %>
var ahora = new Date();
<% Else %>
var ahora = new Date(<%= year(l_tesfec) & "," & month(l_tesfec) & "," & day(l_tesfec) & "," & split(l_teshor,":")(0) & "," &  split(l_teshor,":")(1) & ",0,0"%>);
<% End If %>

function showtime() {

	var timerID;
	var today = new Date();
	var startday = new Date();
	var secPerDay = 0;
	var minPerDay = 0;
	var hourPerDay = 0;
	var secsLeft = 0;
	var secsRound = 0;
	var secsRemain = 0;
	var minLeft = 0;
	var minRound = 0;
	var minRemain = 0;
	var timeRemain = 0;	

	startday = ahora;
	//startday.setYear(<%= year(now)%>);
	today = new Date();

	secsPerDay = 1000 ;
	minPerDay = 60 * 1000 ;
	hoursPerDay = 60 * 60 * 1000;
	PerDay = 24 * 60 * 60 * 1000;
	secsLeft = (today.getTime() - startday.getTime()) / minPerDay;
	secsRound = Math.round(secsLeft);
	secsRemain = secsLeft - secsRound;
	secsRemain = (secsRemain < 0) ? secsRemain = 60 - ((secsRound - secsLeft) * 60) : secsRemain = (secsLeft - secsRound) * 60;
	secsRemain = Math.round(secsRemain);
	minLeft = ((today.getTime() - startday.getTime()) / hoursPerDay);
	minRound = Math.round(minLeft);
	minRemain = minLeft - minRound;
	minRemain = (minRemain < 0) ? minRemain = 60 - ((minRound - minLeft) * 60) : minRemain = ((minLeft - minRound) * 60);
	minRemain = Math.round(minRemain - 0.495);
	hoursLeft = ((today.getTime() - startday.getTime()) / PerDay);
	hoursRound = Math.round(hoursLeft);
	hoursRemain = hoursLeft - hoursRound;
	hoursRemain = (hoursRemain < 0) ? hoursRemain = 24 - ((hoursRound - hoursLeft) * 24)  : hoursRemain = ((hoursLeft - hoursRound) * 24);
	hoursRemain = Math.round(hoursRemain - 0.5);

/*	daysLeft = ((today.getTime() - startday.getTime()) / PerDay);
	daysLeft = (daysLeft - 0.5);
	daysRound = Math.round(daysLeft);
	daysRemain = daysRound;
*/
	var timeValue = "" + hoursRemain;
	timeValue += ((minRemain < 10) ? ":0" : ":") + minRemain;
	timeValue += ((secsRemain < 10) ? ":0" : ":") + secsRemain;
	//timeValue = now.getDate();

	document.all.tiempo.value = timeValue;

	timerID = setTimeout("showtime()",1000);
	timerRunning = true;

}

function finalizar(){
	grabar();
	if (confirm('¿ Desea finalizar la evaluación ?') == true){
		document.all.valida.src = "ag_evaluar_eventos_cap_05.asp?tesnro=<%= l_tesnro %>&testie=" + document.all.tiempo.value;
		window.location.reload();
	}
}

</script>
</head>
<!--
<form>
<input type=hidden name=modsel>
</form>-->

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onLoad="showtime();">
<input type="hidden" name="prenro" value=0>
<input type="hidden" name="resnro" value=0>
<input type="hidden" name="pretipo" value=0>
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr >
		<td align="left" class="barra" >&nbsp;</td>
		<td nowrap align="right" class="barra">
			&nbsp;&nbsp;&nbsp;
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<table cellpadding="0" cellspacing="0">
				<tr>
					<!--
					<td width="50%">&nbsp;</td>
					<td nowrap align="right"><b>Tipo de test:&nbsp;</b></td>
					<td ><input type="Text" size="50" class="deshabinp" readonly value="<%= l_testdesc %>"></td>
					<td width="50%">&nbsp;</td>				
					-->
					<td width="50%">&nbsp;</td>
					<td nowrap align="right"><b>Formulario:&nbsp;</b></td>
					<td ><input type="Text" size="50" class="deshabinp" readonly value="<%= l_fordesabr %>"></td>
					<td width="50%">&nbsp;</td>
					<td width="50%">&nbsp;</td>
					<td nowrap align="right"><b>Empleado:&nbsp;</b></td>
					<td ><input type="Text" id="emp" size="50" class="deshabinp" readonly value="<%= l_tercero %>"></td>
					<td width="50%">&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right">&nbsp;
			
		</td>
	</tr>
	<tr>
		<td align="left" class="barmenu">
			<b>Pregunta:</b>
			<a href="javascript:PrevPre();"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Anterior" border="0"></a>	
			<input type="Text" name="pregunta" size="3" class="hidden" value="1" readonly style="text-align: right; vertical-align: bottom;">
			<input type="Text" size="2" class="hidden" value="de" readonly style="text-align: center; vertical-align: bottom;">
			<input type="Text" name="totpregunta" size="3" class="hidden" value="1" readonly style="text-align: left; vertical-align: bottom;">
			<a href="javascript:NextPre();"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Siguiente" border="0"></a>
		</td>
		<td align="right" class="barmenu" nowrap="nowrap">
			<% If l_tesfin = 0 then%>
				<b>Fecha: </b>
				<%= date() %> - 
				<b>Tiempo Evaluación: </b>
				<input type="text" id="tiempo"  class="reloj" value="" size="5" readonly>
				&nbsp;&nbsp;
			<% Else %>
				<b>Tiempo Evaluación: </b>
				<input type="hidden" id="tiempo"  class="reloj" value="" size="5" readonly>
				<input type="text" id="tiempototal"  class="reloj" value="<%= l_testie %>" size="5" readonly>
				&nbsp;&nbsp;
			<% End If %>
			<% If l_tesfin = 0 then%>
				<a href="#" onClick="Javascript:finalizar();"><b>Finalizar</b></a>
				&nbsp;&nbsp;&nbsp;
			<% End If %>
		</td>
	</tr>	
	<% If Cint(l_tesfin) = -1 Then %>	
	<tr valign="top" height="50%">
	<% Else %>	
	<tr valign="top" height="100%">		
	<% End If %>	
		<td colspan="2" style="">
			<iframe name="ifrm" src="ag_evaluar_eventos_cap_03.asp?&pregunta=1&tesnro=<%= l_tesnro%>&ternro=<%= l_ternro %>&fornro=<%= l_fornro %>&tesfin=<%= l_tesfin %>" width="100%" height="100%"></iframe> 
		</td>
	</tr>
	<% If Cint(l_tesfin) = -1 Then %>	
	<tr valign="top" height="50%">
		<td colspan="2" style="">
			<iframe name="ifrm2" src="ag_evaluar_eventos_cap_07.asp?tesnro=<%= l_tesnro%>&ternro=<%= l_ternro %>&fornro=<%= l_fornro %>&tesfin=<%= l_tesfin %>" width="100%" height="100%"></iframe> 
		</td>
	</tr>		
	<% End If %>	
	
	<tr>
		<td colspan="2" height="20" class="barra" align="right">
		<a class=sidebtnABM href="Javascript:Aceptar();">Aceptar</a>
		<!-- <a class=sidebtnABM href="Javascript:window.close();">Cancelar</a> -->
		<iframe name="valida" width="0" height="50%" src="blanc.asp" ></iframe>
		</td>
	</tr>
</table>
</body>
</html>
