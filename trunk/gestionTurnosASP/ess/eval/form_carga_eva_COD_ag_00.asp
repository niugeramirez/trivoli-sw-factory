<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<% 
'=====================================================================================
'Archivo  : form_carga_eva_ag_COD_00.asp
'Objetivo : form de evaluacion desde autogestion para Codelco
'Fecha	  : 03-02-2005 
'Autor	  : CCRossi
'Modif.	  : 
'=====================================================================================

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

Dim l_RUT

Dim l_fecha 
Dim l_ternro
Dim l_revisor
Dim l_evaperdesde
Dim l_evaperhasta

Dim l_evaevedesabr

Dim siguiente
Dim Anterior

Dim l_yaestaba
Dim l_llamadora
l_yaestaba = false

l_evaevenro = request.querystring("evaevenro")
l_empleg    = request.querystring("empleg")
l_revisor   = request.querystring("revisor")

dim l_logeadoempleg
l_logeadoempleg = request.querystring("logeadoempleg")
session("empleg")=l_logeadoempleg

' si viene de consulta de formularios aprobados viene con valor "consulta", sino viene vacio
l_llamadora   = request.querystring("llamadora") 

if l_empleg = "" then
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT min(empleg) FROM empleado INNER JOIN evacab ON (empleado.ternro=evacab.empleado and evacab.evaevenro="
	l_sql = l_sql & l_evaevenro &")"
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
l_evapernro	  = ""
Set rs9 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaevefdesde,evaevefhasta FROM evaevento"
l_sql = l_sql & " where evaevenro = "& l_evaevenro
rsOpen rs9, cn, l_sql, 0
if not rs9.eof then
	l_evaperdesde = rs9(0)
	l_evaperhasta = rs9(1)
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
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<title>Sistema de Gesti&oacute;n del Desempeño <%if ccodelco<>-1 then%>- RHPro &reg;<%end if%></title>
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
	alert('Supervisado error:'+nro);
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
		alert('Supervisado	incorrecto');
}	
else	{

if (empleg != 0) {	
		document.datos.rternro.value = ternro;
		document.datos.rempleg.value = empleg;
		document.datos.revisor.value = terape + ", " + ternom;
	}
	else{
		alert('Supervisado	incorrecto');
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
	if (logeado==-1)
	{
	if (habilitado!==0) 
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
				//abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
				document.carga.location=etaprogread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value;
			}	
			else
				alert('\nNo se encuentra el Programa de Visualización para la Etapa y la Sección.\n\nEn Configuración, Formularios, Secciones, Etapas podrá\nverificar el nombre de los programas de \ncarga y visualización para cada etapa de la Sección.');
		}		
		else
		{
			if (cargaseccionread!=="")
				if (cargaseccionread!=="*"){
					//abrirVentanaH('evaluador_ingreso_eva_00.asp?evldrnro='+deteval,'',200,100);
					document.carga.location=cargaseccionread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value+'&empleg='+document.datos.empleg.value;
				}
				else
					alert('No se encuentra el Programa de Carga para la Sección.\n\nEn Configuración, Tipo de Sección, podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de Visualización a la Etapa ni al Tipo de Sección.');
		}		
	}
	}
	else
	{
		// el evaluador NO ES el logeado 
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
					document.carga.location=cargaseccionread+'?evldrnro='+deteval+'&evaseccnro='+seccion+'&empleado='+document.datos.ternro.value+'&revisor='+document.datos.rempleg.value+'&evapernro='+document.datos.evapernro.value+'&empleg='+document.datos.empleg.value;
				}
				else
					alert('No se encuentra el Programa de Carga para la Sección.\n\nEn Configuración, Tipo de Sección, podrá verificar el nombre de \nlos programas de carga y visualización para el Tipo de Sección.');
			else
				alert('No se ha asociado un programa de Visualización a la Etapa ni al Tipo de Sección.');
		}		
	}
document.estado.location='estadoseccion_eva_ag_00.asp?evldrnro='+deteval+'&evaseccnro='+seccion+'&logeado='+logeado;

}

function actualizarevaluador(evacabnro, evaseccnro, empleado){
	document.evaluadores.location='form_carga_eva_02.asp?evacabnro='+evacabnro+'&evaseccnro='+evaseccnro+'&ternro='+empleado+'&pantalla=<%=l_pantalla%>&logeadoempleg=<%=l_logeadoempleg%>';
	
}

function buscarrevisor(){
if (isNaN(document.datos.rempleg.value)){
	document.datos.rternro.value = "";
	document.datos.rempleg.value = "";
	alert("El número ingresado para el Supervisor no es correcto.");
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
l_sql = "SELECT empleado.terape, empleado.ternom, empleado.empreporta, empleado.ternro, empleado.empfoto, terfecnac FROM empleado INNER JOIN tercero ON tercero.ternro= empleado.ternro "
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
'Response.Write("<script>alert('ternro de empleado "& l_ternro&"');</script>")
'Response.Write("<script>alert('ternro de revisor "& l_rternro&"');</script>")
'buscar la gerencia -----------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT nrodoc FROM ter_doc "
l_sql = l_sql & " WHERE ternro =" & l_ternro
l_sql = l_sql & "   AND tidnro = 21 "  ' hay que er cual es el RUT
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_RUT = l_rs("nrodoc")
else
	l_RUT = "--"
end if	
l_rs.Close
set l_rs=nothing


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
l_sql = l_sql & " AND   his_estructura.tenro = 6 " 
l_sql = l_sql & " AND   his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_gerencia = l_rs("estrdabr")
else
	l_gerencia = "--"
end if	
l_rs.Close
set l_rs=nothing

'buscar el puesto -----------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
l_sql = l_sql & " AND   his_estructura.tenro = 46 " 
l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
l_sql = l_sql & " ORDER BY his_estructura.htetdesde DESC " 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then	
	l_puesto = l_rs("estrdabr")
	l_htetdesde = l_rs("htetdesde")
else
	l_puesto = "--"
end if	
l_rs.Close
set l_rs=nothing

'Calcular edad -----------------------------------------------------------------
l_edad = DateDiff("d",l_terfecnac,Date())
l_edad = year(date()) - year(l_terfecnac) 
if (month(date()) < month(l_terfecnac)) then
	l_edad = l_edad -1
end if	

dim l_texto
l_texto="--"
' Calcular antiguedad en la empresa


'Calcular antig en el puesto -----------------------------------------------------------------
l_antiguedadpuesto = "--"
if trim(l_htetdesde) <> "" and not isnull(l_htetdesde) then
	l_dias = DateDiff("d",l_htetdesde, date())
	l_meses = DateDiff("m",l_htetdesde, date())
	l_anios = DateDiff("yyyy",l_htetdesde, date())
	if cint(l_dias) > 364 then
		if cint(l_meses) > 12 then  
			l_anios = Int(l_meses / 12 )
			l_meses = CInt(cint(l_meses) - cint(l_anios * 12 ))
		end if	
	else
		l_anios = 0
		l_meses = Int(l_dias / 30.5 )
		l_dias  = Int(l_dias - (l_meses * (30.5)))
	end if
	if trim(l_anios)="0" then
	l_antiguedadpuesto = l_meses &" mes/es " & l_dias &" dia/as "
	else
	l_antiguedadpuesto = l_anios &" año/s "& l_meses &" mes/es "
	end if
end if

' Siguiente/Anterior
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT empleg FROM empleado "
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
l_sql = "SELECT  empleg FROM empleado "
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

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT distinct evadetevldor.evaluador, evacab.evacabnro, evacab.cabaprobada FROM evacab inner join evadetevldor  ON evadetevldor.evacabnro=evacab.evacabnro inner join empleado  ON evacab.empleado= empleado.ternro "
l_sql = l_sql & " WHERE evacab.empleado = " & l_teraux 
l_sql = l_sql & " and evacab.evaevenro="& l_evaevenro
l_sql = l_sql & " and evadetevldor.evatevnro = " & cevaluador
'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_rternro   = l_rs("evaluador")
	l_yaestaba=true
end if 
l_rs.Close
set l_rs=nothing

if trim(l_rternro)<>"" and not isnull(l_rternro) and trim(l_rternro)<>"0" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT empleg, terape, ternom, ternro FROM empleado "
	l_sql = l_sql & " WHERE ternro=" & l_rternro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_revisor= l_rs("empleg")
		l_rterape = l_rs("terape")
		l_rternom = l_rs("ternom")
		l_yaestaba=true
	end if
	l_rs.Close
	set l_rs=nothing

	'buscar el puesto -----------------------------------------------------------------
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
	l_sql = l_sql & " WHERE his_estructura.ternro=" & l_rternro
	l_sql = l_sql & " AND   his_estructura.tenro = 46 " 
	l_sql = l_sql & " AND   his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_evalpuesto   = l_rs("estrdabr")
	else
		l_evalpuesto = "--"
	end if	
	l_rs.Close
	set l_rs=nothing
else
' si no tiene empreporta busco el evaluador
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT distinct revisor.empleg, revisor.terape, revisor.ternom, revisor.ternro FROM evacab inner join evadetevldor on evadetevldor.evacabnro = evacab.evacabnro"
	l_sql = l_sql & "	     and evadetevldor.evatevnro="& cevaluador
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
'#6666CC',   // #6666CC mouseover_background_color
'#ffffff',   // #ffffff border_color
'#ffffff',    // #ffffff separator_color
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

HM_Array2 = [
[100,      // menu width
"mouse_x_position",
"mouse_y_position",
jsfont_color,   // font_color
jsmouseover_font_color,   // mouseover_font_color
'EF8B89',   // navy background_color
'#E85855',   // '#6666CC' mouseover_background_color
'#ffffff',   // '#ffffff' border_color
'#ffffff',    // '#ffffff' separator_color
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
["Mi PDI","Javascript:abrirVentana('<%=cdir1%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
["Convenios de Desempeño","Javascript:abrirVentana('<%=cdir2%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
["Tablero de Control","Javascript:abrirVentana('<%=cdir3%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
["Competencias","Javascript:abrirVentana('<%=cdir4%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
["Portafolio","Javascript:abrirVentana('<%=cdir5%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
["Actividades de Formación","Javascript:abrirVentana('<%=cdir6%>','Mi PDI',750,390,'toolbar=yes,location=yes,status=yes,directories=no,menubar=yes,scrollbars=yes,resizable=yes');",1,0,0],
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
<table border="0" cellpadding="0" cellspacing="0" height="5%">
	<tr>
       	<td align="left" class="barra">Sistema de Gesti&oacute;n del Desempeño: <%= l_evaevedesabr %></td>
      	<td nowrap colspan="" colspan="2" align="right" class="barra" valign="middle">
      		<a class=sidebtnSHW href="#" onClick="MenuPopUp('elMenu2',event)" onMouseOut="MenuPopDown('elMenu1')">Información Complementaria</a>
      		<a class=sidebtnSHW href="#" onClick="Javascript:abrirVentana('rep_formulario_eva_ag_00.asp?evaevenro=<%= l_evaevenro %>&ternro=<%=l_ternro%>&titulo=Formulario&llamadora=Auto&logeadoempleg=<%=l_logeadoempleg%>','',750,390);">Imprimir</a>
      		<img align="absmiddle" src="/serviciolocal/shared/images/logo_codelco.gif" border="0">
		</td>
	</tr>

</table>

<table border="0" cellpadding="0" cellspacing="0" height="95%" width="80%">
<tr height="5%">
	<td colspan="3">
		<table  border="0" cellpadding="0" cellspacing="0">
		<tr>
		    <td nowrap><font <%=l_letra%>><b>Supervisado:</b>&nbsp;
		    	<input <%=l_letra%> type="text" disabled value="<%= l_empleg %>" size="8" name="empleg">
				<input <%=l_letra%> style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">
				&nbsp;<font <%=l_letra%>><b>RUT:</b>
				&nbsp;<input <%=l_letra%> class="blanc" type="Text" name="rut" size="8" value="<%=l_rut%>">
				&nbsp;<font <%=l_letra%>><b>Gerencia:</b>
				&nbsp;<input <%=l_letra%> class="blanc" type="Text" name="gerencia" size="25" value="<%=l_gerencia%>">
				&nbsp;<font <%=l_letra%>><b>Cargo:</b>
				&nbsp;<input <%=l_letra%> class="blanc" type="Text" name="puesto" size="25" value="<%=l_puesto%>">
			</td>
		</tr>
		<tr>
			<td nowrap>
				<font <%=l_letra%>><b>Antig. en la Empresa:</b>
				<%dim l_dia, l_mes, l_anio, l_hab
				if trim(l_ternro) <> "" then
					call antigfec (l_ternro, date , l_dia, l_mes, l_anio, l_hab)
					if l_anio = "" or l_anio = 0 then
						if l_mes = "" or l_mes = 0 then
							if l_dia = "" or l_dia = 0 then
							l_texto = " -- "
							else
							l_texto = l_dia & " día/s."
							end if
						else
							l_texto = l_mes & " mes/es " & l_dia & " día/s."
						end if	
					else
						l_texto = l_anio & " año/s " & l_mes & " mes/es " 	
					end if
				end if
				Set cn = Nothing%>  
				 &nbsp;<input <%=l_letra%> class="blanc" type="Text" name="antiguedad" size="25" value='<%=l_texto%>'>
			     &nbsp;<font <%=l_letra%>><b>en el Cargo:</b>
				 &nbsp;<input <%=l_letra%> class="blanc" type="Text" name="antiguedadpuesto" value="<%=l_antiguedadpuesto%>" size="20">
				 &nbsp;<font <%=l_letra%>>-&nbsp;<b>Per&iacute;odo Desde:</b>
				&nbsp;<input <%=l_letra%> class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperdesde %>">
				&nbsp;<font <%=l_letra%>><b>Hasta:</b>
				&nbsp;<input <%=l_letra%> class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperhasta %>">
			</td>
		</tr>
		</table> 
	</td>
</tr>
<tr height="5%">
	<td colspan="6">
	<table  border="0" cellpadding="0" cellspacing="0">
		<tr>
		    <td nowrap align="right"><font <%=l_letra%>><b>Supervisor:</b></td>
			<td colspan="5" nowrap>
			<input class="rev" type="text"  value="<%= l_revisor %>" size="8" name="rempleg" <%if l_yaestaba then %> readonly <%else%> onKeyPress="return Teclarev(event.keyCode)"  
			onchange="javascript:buscarrevisor();"<%end if%>> 
		<%if not l_yaestaba then %>	
			<a onclick="JavaScript:esempleado=false; window.open('help_emp_01.asp','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Supervisado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
				<img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda" border="0">
			</a>
		<%end if%>	
			<input class="rev" style="background : #e0e0de;" readonly type="text" name="revisor" size="35" maxlength="35" value="<%= l_rterape & ", " &l_rternom%>">
			&nbsp <font <%=l_letra%>><b>Cargo:</b><input disabled class="blanc" type="Text" name="evalpuesto" size="30" value="<%=l_evalpuesto%>">
			</td>
		    <td nowrap align="center" >&nbsp;</td>
		</tr>
	</table> 
	</td>
</tr>

<tr height="90%">
	<td colspan="6" width=100%>
		<table border="0" cellpadding="0" cellspacing="0" height="100%">
	        <tr valign="top" height="10%">
	        	<td align="left" style="" width="650">
   	  				<iframe name="secciones" src="blanc.asp" width="100%" height="100"></iframe> 
   				</td>
				<td style="" width="300">
	    	  		<iframe name="evaluadores" src="blanc.asp" width="100%" height="100%"></iframe> 
				</td>
   			</tr>
   			<tr valign="top" height="<%if l_pantalla="1024" then%>10%<%else%>13%<%end if%>">
   				<td align="left" style="" colspan="2" width="100%">
   	  				<iframe name="estado" src="blanc.asp" width="100%" height="100%"></iframe> 
   				</td>
   			</tr>
   			<tr valign="top" height="90%">
   				<td align="left" style="" colspan="2" width="100%">
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
		alert('El Supervisado no esta relacionado al Evento del Ciclo de Gestión del Desempeño.');
<%		end if
end if%>	
</script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>
