<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<% 
'=====================================================================================
'Archivo  : form_carga_eva_DEL_00.asp
'Objetivo : formulario encabezado, con datos de Deloitte.
'Fecha	  : 08-02-2005   
'Autor	  : Leticia Amadio.
' Modificado: 
'=====================================================================================
'on error goto 0
Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_empleg
Dim l_empfoto
Dim l_evaevenro

Dim l_evapernro ' para pasarle a la Objetivos (periodo)

Dim l_rterape 
Dim l_rternom 
Dim l_rempleg
Dim l_rternro
	'---
Dim l_consejero

Dim l_evaproynro 
Dim l_evaclinom 
Dim	l_evaclicodext 
Dim	l_evaengdesabr 
Dim	l_evaengcodext 

Dim l_cabaprobada ' para cambio de marca
Dim l_evldrnro ' para cambio de estados

Dim l_fecha 
Dim l_ternro
Dim l_revisor
Dim l_evaperdesde
Dim l_evaperhasta
Dim l_evacabnro
Dim l_evaevedesabr

Dim siguiente
Dim Anterior

Dim l_yaestaba


l_yaestaba = false
l_evaevenro = request.querystring("evaevenro")
session("empleg")=""
l_empleg = request.querystring("empleg")
l_revisor = request.querystring("revisor")
	
if trim(l_empleg) = "" then
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT min(empleg) FROM empleado "
	l_sql = l_sql & " INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro="& l_evaevenro &")"
	rsOpen rs9, cn, l_sql, 0
	'response.write l_sql
	if not rs9.eof then
		l_empleg = rs9(0)
	end if
	rs9.close
	set rs9=nothing
	
	if l_empleg&"vacio" = "vacio" then
		l_empleg = "0"
	end if
end if

'response.write l_sql 
'response.write l_empleg
'response.end
'response.write("<script>alert('"&l_empleg&"')</script>")
if l_revisor="" then
	l_revisor="0"	
end if

Set rs9 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaperdesde,evaperhasta, evapernro "
l_sql = l_sql & "FROM evaevento INNER JOIN evaperiodo ON evaevento.evaperact = evaperiodo.evapernro "
l_sql = l_sql & "where evaevenro = "& l_evaevenro
rsOpen rs9, cn, l_sql, 0
if not rs9.eof then
	l_evapernro	  = rs9("evapernro")
	l_evaperdesde = rs9(0)
	l_evaperhasta = rs9(1)
end if
rs9.close
set rs9=nothing
	
' selecciono  -------- 
l_evaproynro=""
Set rs9 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT   evaengdesabr, evaengcodext, evaclinom, evaclicodext, evaproyecto.evaproynro "
l_sql = l_sql & " FROM evaevento " 
l_sql = l_sql & " INNER JOIN evaproyecto ON evaevento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evaengage ON evaengage.evaengnro = evaproyecto.evaengnro "
l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
l_sql = l_sql & " WHERE evaevenro = " & l_evaevenro
RS9.Maxrecords = 1 
rsOpen rs9, cn, l_sql, 0 
if not rs9.eof then 
	l_evaclinom = rs9("evaclinom")
	l_evaclicodext = rs9("evaclicodext")
	l_evaengdesabr = rs9("evaengdesabr")
	l_evaengcodext = rs9("evaengcodext")
	l_evaproynro   = rs9("evaproynro")
end if 
rs9.close
set rs9=nothing

dim l_letra
dim l_pantalla
l_pantalla = request("pantalla")
if l_pantalla = "1024" then
	l_letra="style=font-size:8pt font-type:tahoma"
else	
	l_letra="style=font-size:7pt font-type:arial"
end if
%>
<html>
<head>

<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Sistema de Gesti&oacute;n de Desempeño - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
var cargaseccion = "";
var cargaseccionread = "";


function recargar(){
window.location.reload();
}

function Sig_Fecha(param)
{
pepe = document.datos.fecha.value
pepe = pepe.substr(3, 2) + "/" + pepe.substr(0, 2) + "/" + pepe.substr(6, 4);  
fecha1 = new Date(pepe).valueOf()
pepe2 = 1 * 24 * 60 * 60 * 1000
if (param == 1)
	fecha3 = new Date(fecha1 + pepe2)
else
	fecha3 = new Date(fecha1 - pepe2)
mes = fecha3.getMonth() + 1
mes = mes.toString();
if (mes.length < 2)
	{ 
   	mes = "0" + mes;
	}
dia = fecha3.getDate()
dia = dia.toString();
if (dia.length < 2)
	{ 
   	dia = "0" + dia;
	}
document.datos.fecha.value = dia + "/" + mes + "/" + fecha3.getYear()
document.location ="tabler_emp_gti_00.asp?empleg=" + document.datos.empleg.value + "&fecha=" + document.datos.fecha.value;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Sig_Ant(leg)
{
if (leg != "")
	{
	document.location ="form_carga_eva_DEL_00.asp?evaevenro=<%= l_evaevenro %>&empleg=" + leg + "&revisor=" + document.datos.rempleg.value+ "&pantalla=<%=l_pantalla%>" ;
	}
}

function Volver_primero()
{
	document.location ="form_carga_eva_DEL_00.asp?evaevenro=<%= l_evaevenro %>&pantalla=<%=l_pantalla%>";
}


function Tecla(num){
  if (num==13) {
		verificacodigo(document.datos.empleg,document.datos.empleado,'empleg','terape, ternom','empleado');
		Sig_Ant(document.datos.empleg.value);
		return false;
  }
  return num;
}

function Teclarev(num){
  if (num==13) {
  		buscarrevisor();
		return false;
  }
  return num;
}

function emplerror(nro){
	alert('empleado error:'+nro);
}


function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}
var esempleado=true;
function nuevoempleado(ternro,empleg,terape,ternom)
{

if (esempleado){
	if (empleg != 0) {	
		document.datos.ternro.value = ternro;
		document.datos.empleg.value = empleg;
		document.datos.empleado.value = terape + ", " + ternom;
		Sig_Ant(document.datos.empleg.value);
	}
	else
		alert('Empleado	incorrecto');
}	
else	{

if (empleg != 0) {	
		document.datos.rternro.value = ternro;
		document.datos.rempleg.value = empleg;
		document.datos.revisor.value = terape + ", " + ternom;
	}
	else{
		alert('Empleado	incorrecto');
		document.datos.rternro.value = "";
		document.datos.rempleg.value = "";
		document.datos.revisor.value = "";
	}	
}	
esempleado=true;
}

function actualizarcarga(deteval,seccion,habilitado,etaprogcarga,etaprogread,aprobada,logeado){
	// Primero se busca programa asociado a la etapa y luego al tipo
	// buscar programa de carga
	
	
	if ((habilitado!==0) || (aprobada==0))
	{
		
		if (etaprogcarga!=="")
		{
			if (etaprogcarga!=="*"){
				abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
				document.carga.location=etaprogcarga+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
			}	
			else
				alert('\nNo se encuentra el Programa de Carga para la Etapa y la Sección.\n\nEn Configuración, Formularios, Secciones, Etapas \npodrá verificar el nombre de los programas de \ncarga y visualización para cada etapa de la Sección.');
		}
		else
		{
			if (cargaseccion!=="")
				if (cargaseccion!=="*"){
					abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccion+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
				}
				else
					alert('\nNo se encuentra el Programa de Carga para la Sección.\n\nEn Configuración, Tipo de Sección podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de carga a la Etapa ni al Tipo de Sección.');
		}		
	}
	// buscar programa de visualizacion
	else
	{
		if (etaprogread!==""){
			if (etaprogread!=="*"){
				abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
				document.carga.location=etaprogread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
			}	
			else
				alert('\nNo se encuentra el Programa de Visualización para la Etapa y la Sección.\n\nEn Configuración, Formularios, Secciones, Etapas podrá\nverificar el nombre de los programas de \ncarga y visualización para cada etapa de la Sección.');
		}		
		else
		{
			if (cargaseccionread!=="")
				if (cargaseccionread!=="*"){
					abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccionread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
				}
				else
					alert('No se encuentra el Programa de Visualización para la Sección.\n\nEn Configuración, Tipo de Sección, podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de Visualización a la Etapa ni al Tipo de Sección.');
		}		
	}	
	
	document.datos.all.evldrnro.value=deteval;	

document.estado.location='estadoseccion_eva_00.asp?evldrnro='+deteval+'&evaseccnro='+seccion;	
}

function evldrnro(){
	abrirVentana('cambio_estado_evldrnro_eva_00.asp?evldrnro='+document.datos.all.evldrnro.value,'',350,200);
}

function buscarrevisor(){
if (isNaN(document.datos.rempleg.value)){
	document.datos.rternro.value = "";
	document.datos.rempleg.value = "";
	alert("El legajo ingresado para el Revisor no es correcto."); // consejero ..
	}
else {
	esempleado=false;
	abrirVentanaH('nuevo_emp.asp?empleg='+document.datos.rempleg.value,'',200,100)
	}
}
var Hcarga = 305

function resizecarga(i){
if (i==1){
	Hcarga = Hcarga + 10;
	document.all.carga.style.height = Hcarga;
}
else{	
	Hcarga = Hcarga - 10;
	document.all.carga.style.height = Hcarga;
}	
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


<% '== B O D Y =========================================================================

dim l_gerencia
dim l_puesto
dim l_rut
dim l_terfecnac
dim l_htetdesde 
dim l_antiguedadpuesto
dim l_evalpuesto

dim l_calculo
dim l_dias 
dim l_meses 
dim l_anios 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaevedesabr  FROM  evaevento WHERE evaevenro=" & l_evaevenro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_evaevedesabr = l_rs("evaevedesabr")
end if
l_rs.Close
set l_rs=nothing

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleado.terape, empleado.ternom, empleado.ternro, empleado.empfoto, terfecnac "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " INNER JOIN tercero ON tercero.ternro= empleado.ternro "
l_sql = l_sql & " WHERE empleg=" & l_empleg
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_terape = l_rs("terape")
	l_ternom = l_rs("ternom")
	l_ternro = l_rs("ternro")
	l_empfoto = l_rs("empfoto")
	l_terfecnac= l_rs("terfecnac")
else
	l_ternro=""
end if	
l_rs.Close
set l_rs=nothing

if trim(l_ternro)<>"" then 
	
end if

'Calcular edad -----------------------------------------------------------------
'Calcular antig en el puesto -----------------------------------------------------------------


' Siguiente/Anterior
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleg "
l_sql = l_sql & "FROM empleado "
l_sql = l_sql & "INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro=" & l_evaevenro & " ) "
l_sql = l_sql & "where empleg < " & l_empleg & " ORDER BY empleg DESC"
l_sql = fsql_first (l_sql,1)
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	anterior = l_rs("empleg")
else
	anterior = l_empleg
end if
l_rs.Close

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  empleg "
l_sql = l_sql & "FROM empleado "
l_sql = l_sql & "INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro=" & l_evaevenro & " ) "
l_sql = l_sql & "where empleg > " & l_empleg & " ORDER BY empleg ASC"
l_sql = fsql_first (l_sql,1)
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	siguiente = l_rs("empleg")
else
	siguiente = l_empleg
end if
l_rs.Close
set l_rs=nothing

Dim l_teraux
if l_ternro = "" then
	l_teraux = "0"
else
	l_teraux = l_ternro
end if
'response.write("<script>alert('teraux= "&l_teraux&"');</script>")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'l_sql = "SELECT distinct empleado.empleg, empleado.empreporta, evacab.evacabnro, evacab.cabaprobada "
l_sql = "SELECT distinct evadetevldor.evaluador, evacab.evacabnro, evacab.cabaprobada "
l_sql = l_sql & "FROM evacab "
l_sql = l_sql & " inner join evadetevldor  ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " inner join empleado  ON evacab.empleado= empleado.ternro "
l_sql = l_sql & " WHERE evacab.empleado =" & l_teraux 
l_sql = l_sql & " and evacab.evaevenro="& l_evaevenro
l_sql = l_sql & " and ( evadetevldor.evatevnro = " & cconsejero & " OR  evadetevldor.evatevnro =" & cevaluador & ") "
' response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_evacabnro= l_rs("evacabnro")
	l_cabaprobada = l_rs("cabaprobada")
	l_consejero = l_rs("evaluador")
	l_rternro   = l_rs("evaluador")
	l_yaestaba=true
end if 
l_rs.Close
set l_rs=nothing


'response.write("<script>alert('ternro del Revisor (empreporta:) = "&l_rternro&"');</script>")

if (trim(l_rternro)<>"" and not isnull(l_rternro) and trim(l_rternro)<>"0") then
	'Set l_rs = Server.CreateObject("ADODB.RecordSet")
	'l_sql = "SELECT empleg, terape, ternom, ternro "
	'l_sql = l_sql & "FROM empleado "
	'l_sql = l_sql & "WHERE ternro=" & l_rternro
	'rsOpen l_rs, cn, l_sql, 0
	'if not l_rs.eof then
		'l_revisor= l_rs("empleg")
		'l_rterape = l_rs("terape")
		'l_rternom = l_rs("ternom")
		'l_yaestaba=true
	'end if
	'l_rs.Close
	'set l_rs=nothing
	'buscar el puesto -----------------------------------------------------------------
	'Set l_rs = Server.CreateObject("ADODB.RecordSet")
	'l_sql = "SELECT estrdabr, htetdesde  "
	'l_sql = l_sql & " FROM his_estructura "
	'l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
	'l_sql = l_sql & " WHERE his_estructura.ternro=" & l_rternro
	'l_sql = l_sql & " AND   his_estructura.tenro = 4 " 
	'l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
	'l_sql = l_sql & " ORDER BY his_estructura.htetdesde DESC " 
	'rsOpen l_rs, cn, l_sql, 0
	'if not l_rs.eof then	
		'l_evalpuesto   = l_rs("estrdabr")
	'else
		'l_evalpuesto = "--"
	'end if	
	'l_rs.Close
	'set l_rs=nothing
else
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	'response.write("<script>alert('El Supervisado no tiene Supervisor Asignado');</script>")
end if

		'  si  cdeloitte = -1 
		'   si no tiene empreporta busco el evaluado
if l_ternro <> ""  then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT distinct revisor.empleg, revisor.terape, revisor.ternom, revisor.ternro "
	l_sql = l_sql & " FROM evacab "
	l_sql = l_sql & " inner join evadetevldor   on evadetevldor.evacabnro  = evacab.evacabnro "
	l_sql = l_sql & "	    AND ( evadetevldor.evatevnro="& cconsejero & " OR evadetevldor.evatevnro=" & cevaluador & ") "
	l_sql = l_sql & " inner join empleado revisor on revisor.ternro= evadetevldor.evaluador "
	l_sql = l_sql & " WHERE evacab.empleado = " & l_ternro 
	l_sql = l_sql & " and evacab.evaevenro="& l_evaevenro
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_revisor= l_rs("empleg")
		l_rternro= l_rs("ternro")
		l_rterape = l_rs("terape")
		l_rternom = l_rs("ternom")
		l_yaestaba=true
	end if 
	l_rs.Close
	set l_rs=nothing
end if 
'response.write("<script>alert('l_revisor = "&l_revisor&"');</script>")
%>
<script>
function agregar(){
if (document.all.btnagregar.className == "sidebtnABM"){
	if (document.datos.ternro.value == "")
		alert("Debe ingresar un Empleado.");
	else
	{
		//window.resizeTo(800,580);
		top.window.moveTo(0,0);
		if (document.all) {
			top.window.resizeTo(screen.availWidth,screen.availHeight);
		}
		else
		if (document.layers||document.getElementById) {
			if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth)
			{
				top.window.outerHeight = screen.availHeight;
				top.window.outerWidth = screen.availWidth;
			}
		}

		document.carga.location='form_carga_eva_03.asp?evaevenro=<%= l_evaevenro %>&ternro=<%= l_ternro %>';
		document.all.btnagregar.className = "sidebtnDSB";
		document.all.btnborrar.className = "sidebtnABM";
	}	
}
}

function borrar(){
 if (document.all.btnborrar.className == "sidebtnABM"){
	abrirVentanaH('borra_evaluacion_00.asp?evaevenro=<%=l_evaevenro%>&empleado=<%=l_ternro%>','',5,5);
   }
}

function cambiar(){
	abrirVentana('cambio_aprobacion_cabecera_eva_00.asp?evacabnro=<%= l_evacabnro %>','',250,200);
}

function cargardatos(){
	document.secciones.location="form_carga_eva_01.asp?ternro=<%= l_ternro %>&evaevenro=<%=l_evaevenro%>&revisor=<%=l_revisor%>&pantalla=<%=l_pantalla%>";	
}

</script>
</head>
<body  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" width="95%">

<form name="datos" action="" method="post">
<input type="Hidden" name="ternro"		value="<%= l_ternro %>">
<input type="Hidden" name="rternro"		value="<%= l_rternro %>">
<input type="Hidden" name="evapernro"	value="<%= l_evapernro %>">
<input type="Hidden" name="evldrnro"	value="<%= l_evldrnro %>">
<%
Const lngAlcanGrupo = 2
dim salir 

%>
<table border="0" cellpadding="0" cellspacing="0" height="5%">
<tr >
       	<th align="left">Sistema de Gesti&oacute;n de Desempeño: <%= l_evaevedesabr %></th>
      	<th align="right" colspan=2 valign="middle">
			<!-- SACAR __ -->
      		<%if trim(l_evacabnro)<>"" and cejemplo<>-1 and cint(cdeloitte)<>-1 and ccodelco<>-1 then%>
      		<a class=sidebtnSHW href="Javascript:abrirVentana('etapa_cabecera_eva_00.asp?evacabnro=<%= l_evacabnro %>','',250,200);">Cambiar Etapa</a>
      		<%end if%>
																																					<!--  Supervisados a Cargo -->
			<a class=sidebtnSHW href="Javascript:abrirVentana('form_carga_eva_06.asp?evaevenro=<%= l_evaevenro %>&evaevedesabr=<%= l_evaevedesabr %>','',450,400);">Empleados en el Evento</a>
			<a class=sidebtnSHW href="Javascript:abrirVentana('help_emp_01.asp?empleado=empleado','',600,400);">Buscar</a>
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</th>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" height="95%" width="80%">
<tr height="5%">
	<td width="60%">
		<table  border="0" cellpadding="0" cellspacing="0" height="100%">
		<tr>
		    <td align="left"><font <%=l_letra%>><b>Empleado:</b></td>
			<td >
				<a href="JavaScript:Sig_Ant(<%= anterior %>)"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Número Anterior (<%= anterior %>)" border="0"></a>
				<input <%=l_letra%> type="text" onKeyPress="return Tecla(event.keyCode)" value="<%= l_empleg %>" size="8" name="empleg" onchange="javascript:verificacodigo(document.datos.empleg,document.datos.empleado,'empleg','terape, ternom','empleado');Sig_Ant(this.value);">
				<a href="JavaScript:Sig_Ant(<%= siguiente %>)"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Número Siguiente (<%= siguiente %>)" border="0"></a>
				<a onclick="JavaScript:Javascript:abrirVentana('help_emp_01.asp?empleado=empleado','',600,400);" onmouseover="window.status='Buscar Empleado del Evento'" onmouseout="window.status=' '" style="cursor:hand;">
				<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Empleados en el Evento" border="0"></a>
				<font <%=l_letra%>><input <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">
			</td>
		</tr>
		
		<tr>
			<!-- SACAR -->
			<td>
			<font <%=l_letra%>><b>Revisor:</b> </td>
			<td>
				<font <%=l_letra%>>
				<input class="rev" type="text" value="<%= l_revisor %>" size="6" name="rempleg" <%if l_yaestaba then %> readonly <%else%> onKeyPress="return Teclarev(event.keyCode)"  
				onchange="javascript:buscarrevisor();"<%end if%>> 
				<%if not l_yaestaba then %>	
				<a onclick="JavaScript:esempleado=false; window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Empleado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
				<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Empleados" border="0">
				</a>
				<%end if%>	
				<font <%=l_letra%>><input class="rev" style="background : #e0e0de;" readonly type="text" name="revisor" size="30" maxlength="30" value="<%= l_rterape & ", " &l_rternom%>">
				&nbsp;&nbsp;
			 </td>
		</tr>

		</table> 
	</td>
	<td width="5%">&nbsp;&nbsp; &nbsp;&nbsp;</td>
	<td>
		<table border="0" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left" colspan="2">
				<font <%=l_letra%>><b>Per&iacute;odo Desde:</b>&nbsp;
			 	<input <%=l_letra%> class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperdesde %>" 
			 	<font <%=l_letra%>><b>&nbsp;Hasta:</b> 
				<input <%=l_letra%> class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperhasta %>">
			</td>
		</tr>	
		<tr>	
			<% if  l_evaproynro = "" then %>
				<td nowrap>	&nbsp;</td>
				<td nowrap>&nbsp; </td>
			<% else %>
				<td nowrap>	<font <%=l_letra%>><b> Engagement:</b> &nbsp;</td>
				<td nowrap>
					<input class="rev" <%=l_letra%> type="text" disabled value="<%= l_evaengcodext %>" size="8" name="empleg">
					<input class="rev" <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaengdesabr %>">
				</td>
			<% end if%>
		</tr>	
		<tr>
			<% if  l_evaproynro = "" then %>
				<td nowrap>	&nbsp;</td>
				<td nowrap>&nbsp; </td>
			<% else %>	
				<td nowrap width="12%">	<b> Cliente:</b>&nbsp;</b>
				<td nowrap>
					<input class="rev" <%=l_letra%> type="text" disabled value="<%= l_evaclicodext %>" size="8" name="empleg">
					<input class="rev" <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaclinom%>">&nbsp; &nbsp;&nbsp;
				</td>
			<% end if %>
		</tr>
		</table> 
	</td>
</tr>

<tr height="5%">
	<td colspan="3">
		<table  width="98%" border="0" cellpadding="0" cellspacing="0">
		<tr> 
		    <td align="left">&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;</td>
			<td>&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td align="left">
				<font <%=l_letra%>><b>Proceso Aprobado:</b>&nbsp;
				<input class="rev" style="background : #e0e0de;" readonly type=text name="cabaprobada" size=2 value="<%if l_cabaprobada=-1 then%>SI<%else%>NO<%end if%>">
				&nbsp;
				<a name="btnaprobada" class=sidebtnABM href="Javascript:cambiar();"><b>Cambiar Aprobaci&oacute;n</b></a>
			</td>
		 </tr>
		</table> 
	</td>
</tr>


<tr height="80%">
	<td colspan="3">
		<table border="0" cellpadding="0" cellspacing="0" height="100%">
	        <tr valign="top" height="10%">
	        	<td align="left" style="" width="60%">
   	  				<iframe name="secciones" src="blanc.asp" width="100%" height="100"></iframe> 
   				</td>
  				<td align="left" style="" width="<%if l_pantalla="1024" then%>4%<%else%>5%<%end if%>">
   	  				<iframe name="estado" src="blanc.asp" width="100%" height="100%"></iframe> 
   				</td>
				<td style="" width="35%">
					<a name="btnevldrnro" class=sidebtnABM href="Javascript:evldrnro();"><b>Cambiar Estados</b></a>					
	    	  		<iframe name="evaluadores" src="blanc.asp" width="98%" height="85"></iframe> 
				</td>
   			</tr>
   			<tr valign="top" height="70%">
   				<td align="left" style="" colspan="3" width="100%">
   	  				<iframe name="carga" src="blanc.asp" width="99%" height="100%"></iframe> 
   	  				<iframe name="auxiliar" src="blanc.asp" width="0%" height="0%" style="visibility:hidden"></iframe> 
   				</td>
   			</tr>
   			
		</table>
	</td>
</tr>

</table>

<script>
<%if l_yaestaba then %>
	//window.resizeTo(800,580);
	top.window.moveTo(0,0);
		if (document.all) {
			top.window.resizeTo(screen.availWidth,screen.availHeight);
		}
		else
		if (document.layers||document.getElementById) {
			if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth)
			{
				top.window.outerHeight = screen.availHeight;
				top.window.outerWidth = screen.availWidth;
			}
		}
	cargardatos();
<%else%>	
	window.resizeTo(800,162);
	top.window.moveTo(0,0);
		if (document.all) {
			top.window.resizeTo(screen.availWidth,screen.availHeight);
		}
		else
		if (document.layers||document.getElementById) {
			if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth)
			{
				top.window.outerHeight = screen.availHeight;
				top.window.outerWidth = screen.availWidth;
			}
		}
	<%if l_empleg <> "0" then %>
		alert('El Empleado no esta relacionado al Evento de Evaluación');
	<%end if
end if%>	
</script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>
