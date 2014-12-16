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
'Archivo  : form_carga_eva_Del_00.asp ( es de Deloitte)
'Objetivo : form de evaluacion desde autogestion
'Fecha	  : 06-05-2004 
'Autor	  : CCRossi
'Modif.	  : CCRossi - 25-10-2004 - Mostrar boton AYUDA si cejemplo<>-1 (no es ABN)
'Modif.	  : CCRossi - 03-11-2004 - LABEL Revisor si no es ABN sino Evaluador
'Modif.	  : CCRossi - 03-11-2004 - Agregar reporte Borrador a Boton Impresion
'			Leticia Amadio - 30-12-2004 - Cambio del encabezado para Deloitte 
'			Leticia Amadio - 31-05-2005 - Cambio del encabezado para que vea si es rdp o rde 
'			Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion
'=====================================================================================
on error goto 0
Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_empleg
Dim l_empfoto
Dim l_evaevenro

Dim l_evaproynro 
Dim l_evaclinom  
Dim	l_evaclicodext 
Dim	l_evaengdesabr 
Dim	l_evaengcodext 

Dim l_evapernro ' para pasarle a la Objetivos (periodo)

Dim l_rterape 
Dim l_rternom 
Dim l_rempleg
Dim l_rternro


Dim l_fecha 
Dim l_ternro
Dim l_revisor
Dim l_evaperdesde
Dim l_evaperhasta

Dim l_evaevedesabr

'Dim siguiente
'Dim Anterior

Dim l_yaestaba
Dim l_llamadora
Dim l_evacabnro
l_yaestaba = false

l_evaevenro = request.querystring("evaevenro")
l_empleg    = request.querystring("empleg")
l_revisor   = request.querystring("revisor")

dim l_logeadoempleg
l_logeadoempleg = request.querystring("logeadoempleg")
session("empleg")=l_logeadoempleg

' BUSCAR el cliente, el engagemente del proyecto del evaluado, con la evacabnro
l_evacabnro = request.querystring("evacabnro")


' si viene de consulta de formularios aprobados viene con valor "consulta", sino viene vacio
l_llamadora   = request.querystring("llamadora") 

if l_empleg = "" then
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT min(empleg) FROM empleado "
	l_sql = l_sql & " INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro="& l_evaevenro &")"
	rsOpen rs9, cn, l_sql, 0
	if not rs9.eof then
		l_empleg = rs9(0)
	end if
	rs9.close
	set rs9=nothing
	if l_empleg&"vacio" = "vacio" then
		l_empleg = "0"
	end if
end if

if l_revisor="" then
	l_revisor="0"	
end if

Set rs9 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaperdesde,evaperhasta, evapernro "
l_sql = l_sql & "FROM evaevento INNER JOIN evaperiodo ON evaevento.evaperact = evaperiodo.evapernro "
l_sql = l_sql & "where evaevenro = "& l_evaevenro
RS9.Maxrecords = 1
rsOpen rs9, cn, l_sql, 0
if not rs9.eof then
	l_evapernro	  = rs9("evapernro")
	l_evaperdesde = rs9(0)
	l_evaperhasta = rs9(1)
end if
rs9.close
set rs9=nothing

' selecciono  datos del cliente y engagedel proy -------
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
<title>Formulario de Carga - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>

<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>

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


function Volver_primero()
{
	document.location ="form_carga_eva_00.asp?evaevenro=<%= l_evaevenro %>&pantalla=<%=l_pantalla%>";
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

function actualizarcarga(deteval,seccion,habilitado,etaprogcarga,etaprogread,aprobada,logeado)
{
	// Primero se busca programa asociado a la etapa y luego al tipo
	// buscar programa de carga
	//alert('entra a actualizar'+logeado);
	if (logeado==-1)	{
	if (habilitado!==0) 	{
		if (etaprogcarga!==""){
			if (etaprogcarga!=="*"){
				abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
				document.carga.location=etaprogcarga+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
			}	
			else
				alert('\nNo se encuentra el Programa de Carga para la Etapa y la Sección.\n\nEn Configuración, Formularios, Secciones, Etapas \npodrá verificar el nombre de los programas de \ncarga y visualización para cada etapa de la Sección.');
		}else{
			if (cargaseccion!=="")
				if (cargaseccion!=="*"){
					abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccion+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value+'&empleg='+document.datos.empleg.value;
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
		}else{
			if (cargaseccionread!=="")
				if (cargaseccionread!=="*"){
					abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccionread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value+'&empleg='+document.datos.empleg.value;
				}
				else
					alert('No se encuentra el Programa de Carga para la Sección.\n\nEn Configuración, Tipo de Sección, podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de Visualización a la Etapa ni al Tipo de Sección.');
		}		
	}
	}else	{
		// el evaluador NO ES el logeado 
		if (etaprogread!==""){
			if (etaprogread!=="*"){
				abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
				document.carga.location=etaprogread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
			}	
			else
				alert('\nNo se encuentra el Programa de Visualización para la Etapa y la Sección.\n\nEn Configuración, Formularios, Secciones, Etapas podrá\nverificar el nombre de los programas de \ncarga y visualización para cada etapa de la Sección.');
		}else{
			if (cargaseccionread!=="")
				if (cargaseccionread!=="*"){
					abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccionread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value+'&empleg='+document.datos.empleg.value;
				}
				else
					alert('No se encuentra el Programa de Carga para la Sección.\n\nEn Configuración, Tipo de Sección, podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de Visualización a la Etapa ni al Tipo de Sección.');
		}		
	}
	
//document.estado0.location='estadoseccion_eva_DEL_ag_00.asp?evldrnro='+deteval+'&evaseccnro='+seccion+'&logeado='+logeado;

document.estado.location='estadoseccion_eva_DEL_ag_00.asp?evldrnro='+deteval+'&evaseccnro='+seccion+'&logeado='+logeado;
}

function actualizarevaluador(evacabnro, evaseccnro, empleado){
	document.evaluadores.location='form_carga_eva_02.asp?evacabnro='+evacabnro+'&evaseccnro='+evaseccnro+'&ternro='+empleado+'&pantalla=<%=l_pantalla%>&logeadoempleg=<%=l_logeadoempleg%>';
	
}

function buscarrevisor(){
if (isNaN(document.datos.rempleg.value)){
	document.datos.rternro.value = "";
	document.datos.rempleg.value = "";
	alert("El legajo ingresado para el Revisor no es correcto.");
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


<%
dim l_gerencia
dim l_puesto
dim l_edad
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
l_sql = "SELECT empleado.terape, empleado.ternom, empleado.ternro, empleado.empfoto, terfecnac, empleado.empreporta "
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
	l_rternro  = l_rs("empreporta")
else
	l_ternro=""
end if	
l_rs.Close
set l_rs=nothing

Dim l_teraux
if l_ternro = "" then
	l_teraux = "0"
else
	l_teraux = l_ternro
end if

'if (cdeloitte<>-1) and (trim(l_rternro)<>"" and not isnull(l_rternro) and trim(l_rternro)<>"0") then
	'---
'else 

	'if cint(cdeloitte)=-1 and l_ternro<>"" then
	if l_ternro <> ""  then
		' si no tiene empreporta busco el evaluador
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT distinct revisor.empleg, revisor.terape, revisor.ternom, revisor.ternro "
		l_sql = l_sql & " FROM evacab "
		l_sql = l_sql & " inner join evadetevldor   on evadetevldor.evacabnro   = evacab.evacabnro"
		l_sql = l_sql & "	     and ( evadetevldor.evatevnro="& cevaluador & " OR evadetevldor.evatevnro="& cconsejero & " )"
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
'end if

'response.write ("<script>alert('"&l_ternro&"');</script>")
%>
<script>
function cargardatos(){
	document.secciones.location="form_carga_eva_01.asp?ternro=<%= l_ternro %>&evaevenro=<%=l_evaevenro%>&revisor=<%=l_revisor%>&pantalla=<%=l_pantalla%>&logeadoempleg=<%=l_logeadoempleg%>";	
}

HM_Array1 = [
[100,      // menu width
"mouse_x_position",
"mouse_y_position",
jsfont_color,   // font_color
jsmouseover_font_color,   // mouseover_font_color
'navy',   // background_color
'#6666CC',   // mouseover_background_color
'#ffffff',   // border_color
'#ffffff',    // separator_color
0,         // top_is_permanent
0,         // top_is_horizontal
0,         // tree_is_horizontal
1,         // position_under
1,         // top_more_images_visible
1,         // tree_more_images_visible
"null",    // evaluate_upon_tree_show
"null",    // evaluate_upon_tree_hide
0,         // right_to_left
],         // display_on_click
["Formulario Completo","Javascript:abrirVentana('rep_formulario_eva_ag_00.asp?evaevenro=<%= l_evaevenro %>&ternro=<%=l_ternro%>&titulo=Formulario&llamadora=Auto&logeadoempleg=<%=l_logeadoempleg%>','',750,390);",1,0,0],
<%if cejemplo=-1 then%>
["Borrador","Javascript:abrirVentana('rep_formulario_eva_ag_00.asp?evaevenro=<%= l_evaevenro %>&ternro=<%=l_ternro%>&titulo=Borrador&llamadora=Auto&logeadoempleg=<%=l_logeadoempleg%>','',750,390);",1,0,0],
<%end if%>
]

</script>
</head>
<body  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="" method="post">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">
<input type="Hidden" name="rternro" value="<%= l_rternro %>">
<input type="Hidden" name="evapernro" value="<%= l_evapernro %>">
<%
Const lngAlcanGrupo = 2
dim salir 

'cn.Close
'set cn = Nothing

%>
<table border="0" cellpadding="0" cellspacing="0" height="4%">
	<tr>
       	<th align="left" class="th2">Formulario de Carga de : <%= l_evaevedesabr %></th>
      	<td nowrap colspan="" colspan="2" align="right" class="th2" valign="right">
      		<a class=sidebtnSHW href="#" onClick="MenuPopUp('elMenu1',event)" onMouseOut="MenuPopDown('elMenu1')">Imprimir</a>&nbsp;
      		<%if trim(l_llamadora) <> "consulta" then%>
      		<a class=sidebtnSHW href="Javascript:abrirVentana('form_carga_eva_06.asp?evaevenro=<%= l_evaevenro %>&evaevedesabr=<%= l_evaevedesabr %>','',450,400);">Empleados en el Evento</a> &nbsp;
      		<%end if%>
		</td>
	</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" height="96%" width="80%">
<tr height="5%">
	<td colspan="3">
		<table  border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td nowrap><font <%=l_letra%>><b>Empleado:</b>&nbsp;</td>
			<td nowrap>
		    	<input class="rev" <%=l_letra%> type="text" disabled value="<%= l_empleg %>" size="8" name="empleg">
				<input class="rev" <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">
			</td>
			<% if l_evaproynro = "" then %>
			<td nowrap width="12%"><font <%=l_letra%>>	&nbsp;
			<td nowrap> &nbsp; </td>
			<% else %>
			<td nowrap width="12%"><font <%=l_letra%>>	<b> Cliente:</b>&nbsp;
			<td nowrap>
				<input class="rev" <%=l_letra%> type="text" disabled value="<%= l_evaclicodext %>" size="8" name="empleg">
				<input class="rev" <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaclinom%>">
			</td>
			<% end if%>
		</tr>
		<tr>
			<td nowrap align="left"><font <%=l_letra%>><b>Revisor:</b></td>
			<td colspan="1" nowrap>
			<input class="rev" type="text"  <%=l_letra%> value="<%= l_revisor %>" size="8" name="rempleg" <%if l_yaestaba then %> readonly <%else%> onKeyPress="return Teclarev(event.keyCode)"	onchange="javascript:buscarrevisor();"<%end if%> disabled> 
			<%if not l_yaestaba then %>	
			<a onclick="JavaScript:esempleado=false; window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Empleado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
				<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Empleados" border="0">
			</a>
			<%end if%>	
			<input class="rev" style="background : #e0e0de;" readonly type="text" name="revisor" size="35" maxlength="35" value="<%= l_rterape & ", " &l_rternom%>">
			</td>
			<% if l_evaproynro = "" then %>
			<td nowrap width="12%"><font <%=l_letra%>>	&nbsp;
			<td nowrap> &nbsp; </td>
			<% else %>
			<td nowrap>	<font <%=l_letra%>><b> Engagement:</b> &nbsp;</td>
			<td nowrap>
				<input class="rev" <%=l_letra%> type="text" disabled value="<%= l_evaengcodext %>" size="8" name="empleg">
				<input class="rev" <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaengdesabr %>">
			</td>
			<% end if%>
		</tr>
		</table> 
	</td>
</tr>
<tr height="95%">
	<td colspan="3" width=100%>
		<table border="0" cellpadding="0" cellspacing="0" height="100%">
			<tr valign="top">
	        	<td align="left" style="" width="430"> 	</td>
				<td width="30">&nbsp;</td>
				<td style="" width="320"></td>
   			</tr>
			
	        <tr valign="top" height="15%">
	        	<td align="left" style="" width="430">
   	  				<iframe name="secciones" src="blanc.asp" width="100%" height="100"></iframe> 
   				</td>
				<td width="30"> <!-- <iframe name="estado0" src="blanc.asp" width="100%" height="100%"></iframe> -->
				</td>
				<td style="" width="320">
					&nbsp;&nbsp;<font <%=l_letra%>><b>Responsables de la secci&oacute;n</b>&nbsp;	<br>
	    	  		<iframe name="evaluadores" src="blanc.asp" width="100%" height="100%"></iframe>
				</td>
   			</tr>
			
			<tr valign="top" align="center" height="<%if l_pantalla="1024" then%>5%<%else%>6%<%end if%>">
   				<td align="center" style="" colspan="3" width="100%"> 
   	  				<iframe name="estado" src="blanc.asp" width="100%" height="100%"></iframe>  
   				</td>
   			</tr>
			
   			<tr valign="top" height="80%">
   				<td align="left" style="" colspan="3" width="100%"> 
   	  				<iframe name="carga" src="blanc.asp" width="100%" height="100%"></iframe> 
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
<%		end if
end if%>	
</script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>
