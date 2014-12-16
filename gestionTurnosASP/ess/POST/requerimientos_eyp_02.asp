<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% on error goto 0
'---------------------------------------------------------------------------------
'Archivo	: requerimientos_eyp_02.asp
'Descripción: Form de datos de nota
'Autor		: Raul CHinestra
'Fecha		: 30-08-2003
' Modificado  : 12/09/2006 Raul Chinestra - se agregó Requerimientos de Personal en Autogestión   
'				27/11/2006 - Mariano Capriz - Se cambio el titulo de "NOTAS" por "Requerimiento de Personal"
'----------------------------------------------------------------------------------


'Variables
 Dim l_empleg
 
 'Datos del formulario
Dim l_reqpernro
Dim l_reqperdesabr 
Dim l_reqperdesext

Dim l_reqpersolpor
Dim l_reqperrelpor
Dim l_reqperent
Dim l_motprinro
Dim l_motreqnro
Dim l_motreqpri
Dim l_reqpercanper
Dim l_reqperrelfec
Dim l_reqpersolfec
Dim l_reqperfecalt

Dim l_puenro
Dim l_areanro
Dim l_reqperrep
Dim l_reqperremofr
Dim l_tcnro
Dim l_reghornro
Dim l_reqperprires
Dim l_reqperpritar
Dim l_reqperotrtar
Dim l_reqperbenofr

 'Dim l_ternro
 Dim l_tipo
 
 Dim l_rs
 Dim l_rs1
 Dim l_sql

 ' tomar parametros de entrada
 l_tipo   = Request.QueryString("tipo")
 l_reqpernro= Request.QueryString("reqpernro")
 
 l_empleg	= request.Querystring("empleg")   
 'l_ternro = l_ess_ternro
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 select Case l_tipo
	Case "A":
		l_reqpernro  	= ""
		l_reqperdesabr 	= ""
		l_reqperdesext 	= ""
		
		l_reqpersolpor	= l_empleg
		l_reqpersolfec	= ""
		l_reqpercanper	= ""
		l_reqperfecalt	= ""
		l_reqperrelpor	= ""
		l_reqperrelfec	= ""
		l_reqperent 	= ""
		l_motreqpri		= ""
		l_motprinro		= ""
		l_motreqnro		= ""
		
		l_puenro		= 0
		l_reqperrep		= ""
		l_reqperremofr	= 0
		l_reqperprires	= ""
		l_reqperpritar	= ""
		l_reqperotrtar	= ""
		l_reqperbenofr	= ""
		
		
	Case "M":
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM pos_reqpersonal "
		l_sql = l_sql & " WHERE reqpernro = " & l_reqpernro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_reqpernro		= l_rs("reqpernro")
			l_reqperdesabr	= l_rs("reqperdesabr")
			l_reqperdesext	= l_rs("reqperdesext")

			l_reqpersolpor	= l_rs("reqpersolpor")
			l_reqpersolfec	= l_rs("reqpersolfec")
			l_reqpercanper	= l_rs("reqpercanper")
			l_reqperfecalt	= l_rs("reqperfecalt")
			l_reqperrelpor	= l_rs("reqperrelpor")
			l_reqperrelfec	= l_rs("reqperrelfec")
			l_reqperent 	= l_rs("reqperent")
			l_motreqpri		= l_rs("motreqpri")
			l_motprinro		= l_rs("motprinro")
			l_motreqnro		= l_rs("motreqnro")
			
			l_puenro		= l_rs("puenro")
			l_reqperrep		= l_rs("reqperrep")
			IF not isnull(l_rs("reqperremofr"))	then 
				l_reqperremofr = replace(l_rs("reqperremofr"),",",".")
			else l_reqperremofr = "" end if
			l_reqperprires	= l_rs("reqperprires")
			l_reqperpritar	= l_rs("reqperpritar")
			l_reqperotrtar	= l_rs("reqperotrtar")
			l_reqperbenofr	= l_rs("reqperbenofr")
			

		end if
		l_rs.Close
 end select
 
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Requerimientos  - Empleos y Postulantes - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_hora.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script>

function Validar_Formulario(){


if (Trim(document.datos.reqperdesabr.value) == ""){
		alert("Debe Ingresar una Descripción Abreviada.");
		document.datos.reqperdesabr.focus();
		return;
}		

if (document.datos.reqperdesext.value.length > 200){
		alert("La Descripción Extendida no puede superar los 200 caracteres.");
		document.datos.reqperdesext.focus();
		return;
}		

if(document.rel.datossol.empleg.value == ""){
		alert("Debe Seleccionar por quien es relevado.");
		document.rel.datossol.empleg.focus();
		return;
}		

if (document.ent.datossol.empleg.value == "")	{
		alert("Debe Seleccionar quien es el entrevistador.");
		document.ent.datossol.empleg.focus();
		return;
}		


if (!validarfecha(document.datos.reqpersolfec)){
		document.datos.reqpersolfec.focus();
		return;		
}

		
if ((document.datos.reqpercanper.value == "")||(!validanumero(document.datos.reqpercanper,4,0))||((document.datos.reqpercanper.value < 0)) ){
		alert('Debe ingresar la Cantidad de Personas.');
		document.datos.reqpercanper.select();
		document.datos.reqpercanper.focus();
		return;		
}		

if (!validarfecha(document.datos.reqperfecalt)){
		document.datos.reqperfecalt.focus();
		return;		
}		

if (isNaN(document.rel.datossol.ternro.value)){
		alert('Debe seleccionar por quien es Relevado');
		document.rel.datossol.empleg.focus();
		return;		
}		

if (!validarfecha(document.datos.reqperrelfec)){
		document.datos.reqperrelfec.focus();
		return;
}		

if (isNaN(document.ent.datossol.ternro.value)){
		alert('Debe seleccionar un Entrevistador');
		document.ent.datossol.empleg.focus();
		return;		
}		

if (document.datos.motprinro.value == 0){
		alert('Debe seleccionar un Motivo de Prioridad');
		document.datos.motprinro.focus();
		return;		
}		

if (document.datos.motreqnro.value == 0){
		alert('Debe seleccionar un Motivo de Pedido');
		document.datos.motreqnro.focus();
		return;		
}		

if (document.datos.puenro.value == "0"){
		alert('Debe seleccionar un Puesto');
		document.datos.puenro.focus();
		return;
}		

if ((document.rep.datossol.empleg.value == "")||(isNaN(document.rep.datossol.empleg.value))){
		alert('Debe seleccionar a quien Reporta');
		document.rep.datossol.empleg.focus();
		return;
}		

if ((!validanumero(document.datos.reqperremofr,15,4))||(document.datos.reqperremofr.value <= 0)||(document.datos.reqperremofr.value == "")){
		alert('Debe ingresar la Remuneración Ofrecida con 15 Enteros y 4 Decimales');
		document.datos.reqperremofr.select();document.datos.reqperremofr.focus();
		return;
}		

if (document.datos.reqperprires.value.length > 500){
		alert("La Principal Responsabilidad no puede superar los 500 caracteres.");
		document.datos.reqperprires.select();document.datos.reqperprires.focus();
		return;
}		

if (document.datos.reqperpritar.value.length > 500){
		alert("La Principal Tarea no puede superar los 500 caracteres.");
		document.datos.reqperpritar.select();document.datos.reqperpritar.focus();
		return;
}		

if (document.datos.reqperotrtar.value.length > 500){
		alert("Las Otras Tareas no pueden superar los 500 caracteres.");
		document.datos.reqperotrtar.select();document.datos.reqperotrtar.focus();
		return;
}		

if (document.datos.reqperbenofr.value.length > 500){
		alert("Los Beneficios Ofrecidos no pueden superar los 500 caracteres.");
		document.datos.reqperbenofr.select();document.datos.reqperbenofr.focus();
		return;
}		

	document.valida.location = "requerimientos_eyp_09.asp?tipo=<%= l_tipo%>&reqpernro=" + document.datos.reqpernro.value + "&reqperdesext=" + document.datos.reqperdesext.value + "&reqperdesabr=" + document.datos.reqperdesabr.value/* + "&ternro=" + document.datos.ternro.value*/;	
	
}

function nada(tecla){
	if (tecla == 13){
		return false;
	}
	return tecla;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
	return w_in.showModalDialog(pagina,'', 'help:0;status:0;resizable:0;center:1;scroll:0;yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 14);
	if (jsFecha == null) txt.value = ''
	else txt.value = jsFecha;
}

function valido(){
  document.datos.submit();
}

function invalido(texto){
  alert(texto);
}


</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="document.datos.reqperdesabr.focus();">
<form name="datos" action="requerimientos_eyp_03.asp?empleg=<%=l_empleg%>" method="post">
<input type="hidden" name="tipo" value="<%=l_tipo%>">
<!-- <input type="hidden" name="ternro" value="<%'=l_ternro%>"> -->
<input type="hidden" name="reqpernro" value="<%=l_reqpernro%>">
<input type="hidden" name="tnonroant" value="<%'=l_tnonroant%>">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
 <tr style="border-color :CadetBlue;">
 	<!-- 27/11/2006 - MDC ------------------------------------------------------>
	<th align="left" class="th2" colspan="3" height="1">Requerimiento de Personal</th>
	<!-- ---------------- ------------------------------------------------------>
	<th class="th2" align="right">&nbsp;</th>
 </tr>
 
<tr>
	<td align="right"><b>Código:</b></td>
	<td><input type="Text" class="deshabinp" readonly size="4" value="<%= l_reqpernro %>" ></td>
</tr>
<tr>
<td align="right"> <b>Desc. Abreviada:</b></td>
<td colspan="3">
		<input type="text" name="reqperdesabr" size="60" maxlength="50" value="<%= l_reqperdesabr %>">
</td>
</tr>
<tr>
    <td align="right"><b>Desc. Extendida:</b></td>
	<td colspan="3" align="left">
	    <textarea name="reqperdesext" rows="3" cols="45" maxlength="200"><%= l_reqperdesext %></textarea>
	</td>
</tr>
<tr>
	<td align="right"><b>Fecha de Solicitud:</b><td>
	<input  type="text" name="reqpersolfec" size="10" maxlength="10" value="<%= l_reqpersolfec %>" >
		<a href="Javascript:Ayuda_Fecha(document.datos.reqpersolfec);">
			<img src="/serviciolocal/shared/images/cal.gif" border="0">
		</a>
</tr>
<tr>
	<td align="right"><b>Cantidad de Personas:</b></td>
	<td><input type="text" name="reqpercanper" size="4" maxlength="4" value="<%= l_reqpercanper %>"></td>
</tr>
<tr>
	<td align="right"><b>Alta Programada:</b><td>
	<input  type="text" name="reqperfecalt" size="10" maxlength="10" value="<%= l_reqperfecalt %>" >
		<a href="Javascript:Ayuda_Fecha(document.datos.reqperfecalt);">
			<img src="/serviciolocal/shared/images/cal.gif" border="0">
		</a>
</tr>
<tr>
    <td align="right" ><b>Relevada por:</b></td>
	<td>
		<input type="Hidden" name="reqperrelpor" value="">
		<iframe name="rel" frameborder="0" width="100%" height="80" scrolling="No" src="requerimiento_rel_eyp_00.asp?ternro=<%= l_reqperrelpor %>"></iframe>
	</td>
</tr>
<tr>
	<td align="right"><b>Fecha Relevamiento:</b><td>
	<input  type="text" name="reqperrelfec" size="10" maxlength="10" value="<%= l_reqperrelfec %>" >
		<a href="Javascript:Ayuda_Fecha(document.datos.reqperrelfec);">
			<img src="/serviciolocal/shared/images/cal.gif" border="0">
		</a>
	
</tr>
<tr>
    <td align="right" ><b>Entrevistador en linea:</b></td>
	<td>
		<input type="Hidden" name="reqperent" value="">
		<iframe name="ent" frameborder="0" width="100%" height="30" scrolling="No" src="requerimiento_ent_eyp_00.asp?ternro=<%= l_reqperent %>"></iframe>
	</td>
</tr>
<tr>
	<td align="right"><b>Prioridad:</b></td>
	<td>
		<input type="Radio" name="motreqpri" value="1" id="r1"><b>Urgente</b> 
		<input type="Radio" name="motreqpri" value="2" id="r2" checked ><b>media</b> 
		<input type="Radio" name="motreqpri" value="3" id="r3"><b>baja</b>
		<script>
		<% If l_motreqpri <> "" then %>
			document.datos.r<%= l_motreqpri %>.checked= true;
		<% End If %>
		</script>
	</td>
</tr>
<tr>
	<td align="right"><b>Motivo Prioridad:</b></td>
	<td>
		<select name=motprinro size="1"  style="width:382;"  >
		<option value="0"  selected> << Seleccione una Opción >> </option><%
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT motpridesabr, motprinro "
			l_sql = l_sql & " FROM pos_motivopri "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof %>	
				<option value="<%= l_rs("motprinro") %>" > 
				<%= l_rs("motpridesabr") %> (<%=l_rs("motprinro")%>) </option>
				<%	l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<%if l_motprinro = "" or isnull(l_motprinro) then %>
			<script> document.datos.motprinro.value = "0" </script>
		<% Else  %>
			<script> document.datos.motprinro.value = "<%= l_motprinro %>" </script>
		<% End If %>
	</td>	
</tr>
<tr>
	<td  align="right"><b>Motivo Pedido:</b></td>
	<td>
		<select name=motreqnro size="1"  style="width:382;"  >
		<option value="0"  selected> << Seleccione una Opción >> </option><%
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT motreqdesabr, motreqnro "
			l_sql = l_sql & " FROM pos_motivoreq "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof %>	
				<option value="<%= l_rs("motreqnro") %>" > 
				<%= l_rs("motreqdesabr") %> (<%=l_rs("motreqnro")%>) </option>
				<%	l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<%if isnull(l_motreqnro) or l_motreqnro = "" then %>
			<script> document.datos.motreqnro.value = "0" </script>
		<% Else  %>
			<script> document.datos.motreqnro.value = "<%= l_motreqnro %>" </script>
		<% End If %>
	</td>
</tr>
 
	<tr>
    <td align="right" ><b>Puesto:</b></td>
	<td>
		<select name=puenro size="1"  style="width:382;"  >
			<option value="0"  selected> << Seleccione una Opción >> </option><%
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT puenro, puedesc "
			l_sql = l_sql & " FROM puesto "
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof %>	
			<option value="<%= l_rs(0) %>" > 
			<%= l_rs(1) %> (<%=l_rs(0)%>) </option>
		<%			l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
		<%if l_tipo = "A" or isnull(l_puenro) or l_puenro = "" then %>
			<script> document.datos.puenro.value = "0" </script>
		<% Else  %>
			<script> document.datos.puenro.value = "<%= l_puenro %>" </script>
		<% End If %>
	</td>	
</tr>
<tr>
    <td align="right" ><b>Reporta a:</b></td>
	<td>
		<input type="Hidden" name="reqperrep" value="">
		<iframe name="rep" frameborder="0" width="100%" height="30" scrolling="No" src="requerimiento_rep_eyp_00.asp?ternro=<%= l_reqperrep %>"></iframe>
	</td>
</tr>
<tr>
    <td align="right"><b>Remuneración Ofrecida:</b></td>
	<td align="left">
	    <input name="reqperremofr" onkeypress="return nada(event.keyCode);" value="<%= l_reqperremofr %>" type="Text" size="20" maxlength="20">
	</td>
</tr>
<tr>
    <td align="right"><b>Principal Responsabilidad:</b></td>
	<td align="left">
	    <textarea name="reqperprires" rows="2" cols="45" maxlength="200" style="width:382;"><%= l_reqperprires %></textarea>
	</td>
</tr>
<tr>
    <td align="right"><b>Principal Tarea:</b></td>
	<td align="left">
	    <textarea name="reqperpritar" rows="2" cols="45" maxlength="200" style="width:382;"><%= l_reqperpritar %></textarea>
	</td>
</tr>
<tr>
    <td align="right"><b>Otras Tareas:</b></td>
	<td align="left">
	    <textarea name="reqperotrtar" rows="2" cols="45" maxlength="200" style="width:382;"><%= l_reqperotrtar %></textarea>
	</td>
</tr>
<tr>
    <td align="right"><b>Beneficios Ofrecidos:</b></td>
	<td align="left">
	    <textarea name="reqperbenofr" rows="2" cols="45" maxlength="200" style="width:382;"><%= l_reqperbenofr %></textarea>
	</td>
</tr>
 
 
 
 <tr>
    <td align="right" class="th2" colspan="4" height="1">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 		
	</td>
 </tr>
</table>
</form>

</body>
</html>


 

 
