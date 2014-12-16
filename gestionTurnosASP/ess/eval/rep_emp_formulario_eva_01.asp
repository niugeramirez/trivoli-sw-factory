<% Option Explicit %>

<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->

<!--#include virtual="/serviciolocal/shared/inc/resumen_planaccion_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_notas_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_vistos_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_totales_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_resultados_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_resultados_eva_ABN.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_grafico_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_cardinales_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_objetivos_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_objetivos_plan_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_objetivossmart_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_objetivossmart_eva_ABN.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_general_eva_ABN.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_plansmart_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_plansmart_eva_ABN.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_datosadm_eva_00.inc"-->

<!-- CODELCO -->
<!--#include virtual="/serviciolocal/shared/inc/resumen_borrador_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_compromisos_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_cierre_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_cierreEva_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_actividades_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_retroalimentacion_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_evalborrador_eva_COD.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_evaluacion_eva_COD.inc"-->

<!-- DELOITTE -->
<!--#include virtual="/serviciolocal/shared/inc/resumen_resultadosyarea_eva_00.inc"-->
<!-- sacarlo !! --> 	<!-- ''''#include virtual="/serviciolocal/shared/inc/resumen_gralobj_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_calificobj_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_areacom_eva_00.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/resumen_objcom_eva_00.inc"-->

<% 'if cint(cdeloitte) = -1 then   asi no funciona..%>
	<!--#include virtual="/serviciolocal/shared/inc/resumen_areacomRDP_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_datosadmRDP_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_calificobjRDP_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_calificcompRDP_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_calificgralRDP_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_calificcompSI_eva_00.inc"-->
	<!--#include virtual="/serviciolocal/shared/inc/resumen_calificcompSE_eva_00.inc"-->
<% ' end if %>
<!--#include virtual="/serviciolocal/shared/inc/resumen_compxestr_eva_00.inc"-->
<% 
'---------------------------------------------------------------------------------------------
' Modificado: 25-10-2004 - CCRossi - Agregar secciones propias de ABN
' Modificado: 25-10-2004 - CCRossi - Agregar Gerencia y Puesto en la Cabecera   
'			  30-12-2004 - Leticia Amadio - cambio de la cabecera para deloitte 
'			  10-01-2005 - L. Amadio - agregar un inc (resumen)
' 			  04-02-2005 - L. Amadio - encabezado deloitte 
' 			  08-02-2005 - CCROssi   - encabezado Codelco y seccines codelco
' 			  19-05-2005 - CCROssi   - cambiar forma de pasaje de parametros para que no haya problemas con cantidad de empleados
' 			  13-07-2005 - CCROssi   - tomar parametro ternro cuando viene del formulario.
'		esto y l_quien="" significa que solo hay quelistar el formulario del ternro.
' 			  08-08-2005 - CCROssi   - cambiar de lugar una busqueda de datos del evento para 
'						 DELOITTE, despues de haber armado la listernro
'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'		     11-08-2006 - LA. -  sacar v_empleado y dejar empleado
'			 24-08-2006 - LA. - modificaciones para que se vean los margenes al imprimir
'			 15-06-2006 - LA. - agregar resumen de la seccion Def Objs con plan de desarrollo 
'---------------------------------------------------------------------------------------------

on error goto 0

Dim l_llamadora

dim l_terfecnac
dim l_htetdesde 
dim l_antiguedadpuesto
dim l_evalpuesto

dim l_calculo
dim l_dias 
dim l_meses 
dim l_anios 

Dim l_sql
Dim l_rs
Dim l_rssecc
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

Dim l_rterape 
Dim l_rternom 
Dim l_rempleg
Dim l_rternro

Dim l_evacabnro
dim l_estrnro

Dim l_ternro
Dim l_revisor
Dim l_evaperdesde
Dim l_evaperhasta

Dim l_evaevedesabr

dim l_gerencia
dim l_puesto
dim l_sector
dim l_rpuesto

dim l_rut

Dim l_titulo_rep
Dim l_listternro

dim l_tit_rep

dim l_borrador
dim l_super
dim l_logeadoternro
l_borrador=0
l_super=0
l_logeadoternro=""

dim l_linea
dim l_cantidadlineas 
l_cantidadlineas = 47

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


l_evaevenro		= request.querystring("evaevenro")
l_tit_rep       = Request.QueryString("titulo")
l_llamadora		= Request.QueryString("llamadora")
l_logeadoternro = Request.QueryString("logeadoternro")

l_revisor="0"	

Dim l_join
Dim l_quien

l_ternro	= Request.QueryString("ternro")
l_join		= Request.QueryString("join")
l_quien		= Request.QueryString("quien")

if trim(l_quien)="evaluador" and trim(l_ternro)="" then
l_ternro=0
end if
if trim(l_quien)="empleado" and trim(l_ternro)="" then
l_ternro=0
end if

' selecciono datos de cliente y engage  
if cdeloitte=-1 then
	l_evaproynro=""
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT   evaengdesabr, evaengcodext, evaclinom, evaclicodext, evaproyecto.evaproynro  FROM evaevento INNER JOIN evaproyecto ON evaevento.evaproynro = evaproyecto.evaproynro "
	l_sql = l_sql & " INNER JOIN evaengage ON evaengage.evaengnro = evaproyecto.evaengnro INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro WHERE evaevenro = " & l_evaevenro
	rsOpen rs9, cn, l_sql, 0 
	if not rs9.eof then 
		l_evaclinom		= rs9("evaclinom")
		l_evaclicodext	= rs9("evaclicodext")
		l_evaengdesabr	= rs9("evaengdesabr")
		l_evaengcodext	= rs9("evaengcodext")
		l_evaproynro	= rs9("evaproynro")
	end if 
	rs9.close
	set rs9=nothing
end if

l_listternro="0"

'=================================================================================
'l_quien estara vacio si viene desde el FORMULARIO
'l_llamadora=AUTO si viene desde autogestion, ya ea del formulario o desde ESS
'=================================================================================

Set rs9 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT empleado FROM evacab "
l_sql =  l_sql & " INNER JOIN empleado ON evacab.empleado=empleado.ternro"
if trim(l_quien)="evaluador" then
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro"
l_sql = l_sql & " AND  evadetevldor.evaluador=" & l_ternro
l_sql = l_sql & " AND  evadetevldor.evatevnro=" & cevaluador 
end if
if trim(l_logeadoternro)<>"" and trim(l_quien)<>"" then
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro"
l_sql = l_sql & " AND  evadetevldor.evaluador=" & l_logeadoternro
l_sql = l_sql & " AND  evadetevldor.evatevnro=" & cevaluador 
end if
if trim(l_quien)="estructura" and trim(l_join)<> "" then
l_sql = l_sql & " " & TRIM(l_join)
end if
l_sql =  l_sql & " WHERE evacab.evaevenro = " & l_evaevenro
if trim(l_quien)="empleado" and trim(l_ternro)<>"" then
l_sql = l_sql & " AND  evacab.empleado=" & l_ternro
end if
if trim(l_quien)="" and (UCAsE(trim(l_llamadora))="AUTO" and trim(l_ternro)<>"" and trim(l_ternro)<>"0") then
l_sql = l_sql & " AND  evacab.empleado=" & l_ternro
end if
rsOpen rs9, cn, l_sql, 0
'Response.Write l_Sql
'Response.End
do while not rs9.eof 
	l_listternro = l_listternro &"," & rs9("empleado")
	rs9.MoveNext
loop
rs9.Close
set rs9=nothing


Set rs9 = Server.CreateObject("ADODB.RecordSet")
if ccodelco=-1 then
l_sql = "SELECT evaevefdesde,evaevefhasta, evaevedesabr FROM evaevento where evaevenro = "& l_evaevenro
else
l_sql = "SELECT evaperdesde,evaperhasta, evaevedesabr FROM evaevento INNER JOIN evaperiodo ON evaevento.evaperact = evaperiodo.evapernro "
l_sql = l_sql & " INNER JOIN evacab  ON evacab.evaevenro= evaevento.evaevenro AND evacab.empleado IN (" & l_listternro &")"
l_sql = l_sql & " where evaevento.evaevenro = "& l_evaevenro
end if
RS9.Maxrecords = 1
rsOpen rs9, cn, l_sql, 0
if not rs9.eof then
	l_evaperdesde = rs9(0)
	l_evaperhasta = rs9(1)
	l_evaevedesabr = rs9(2)
else
	if cdeloitte=-1 then
		response.write ("<script>alert('El empleado no esta en el Proyecto / Evento,')</script>")
		response.end	
	end if
end if
rs9.Close
set rs9=nothing
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Proceso de Gesti&oacute;n de Desempe&ntilde;o<%if ccodelco<>-1 then%>- RHPro &reg;<%end if%></title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
var cargaseccion = "";

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/bsas/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function cambiartam(obj){
document.all.ifrm.style.heigth= 300;

}

function imprimir(){
	//parent.frames.ifrm.focus();
	window.print();
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
.texto
{
	font-size: 9;
}
</style>


<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rssecc = Server.CreateObject("ADODB.RecordSet")

Const lngAlcanGrupo = 2
dim salir 

%>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos" action="" method="post">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">
<input type="Hidden" name="rternro" value="<%= l_rternro %>">

<table border="1" cellpadding="0" cellspacing="0" style=" width: 100%;">
 
<%
l_linea = l_linea + 1
if trim(l_listternro)="0" then%>
<tr height="5">
	<td colspan="3"><b>No hay <%if ccodelco=-1 then%>supervisados<%else%>evaluados<%end if%> para el filtro.</b></td>
</tr>	

<%end if

Dim l_arrternro
l_arrternro = Split(l_listternro,",")

dim l_i
dim l_tieneobj

For l_i=1 To UBound(l_arrternro) 
	' ----
	l_ternro = l_arrternro(l_i)
	l_rternro=""
	l_anios = 0
	l_meses = 0
	l_dias  = 0
	l_antiguedadpuesto=""
	l_texto=""
	l_dia=0
	l_mes=0
	l_anio=0 
	l_hab=0
	
	l_sql = "SELECT terape, ternom, empleg, empfoto FROM empleado WHERE ternro=" & l_ternro
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		l_terape = l_rs("terape")
		l_ternom = l_rs("ternom")
		l_empleg = l_rs("empleg")
	end if	
	l_rs.Close
	
	l_tieneobj  = 0
	
	l_sql = "SELECT distinct evacab.evacabnro, tieneobj FROM evacab WHERE evacab.empleado = " & l_ternro  & " and evacab.evaevenro="& l_evaevenro
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then
		l_evacabnro = l_rs("evacabnro")
		l_tieneobj  = l_rs("tieneobj")
	end if 
	l_rs.Close
	
	if cint(cdeloitte) = -1 then ' busco como revisor : proyrevisor o consejero!
			
		l_sql = "SELECT distinct revisor.empleg, revisor.terape, revisor.ternom, revisor.ternro FROM evacab inner join evadetevldor on evadetevldor.evacabnro = evacab.evacabnro and (evadetevldor.evatevnro="& cevaluador & " OR evadetevldor.evatevnro="&cconsejero &")"
		l_sql = l_sql & " inner join empleado revisor on revisor.ternro= evadetevldor.evaluador WHERE evacab.empleado = " & l_ternro  & " and evacab.evaevenro="& l_evaevenro
		l_rs.Maxrecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_revisor= l_rs("empleg")
			l_rternro= l_rs("ternro")
			l_rterape = l_rs("terape")
			l_rternom = l_rs("ternom")
			'l_yaestaba=true
		end if 
		l_rs.Close
	else
		if trim(l_rternro)="" or isnull(l_rternro) then
			' si no tiene empreporta busco el evaluador
			l_sql = "SELECT distinct revisor.empleg, revisor.terape, revisor.ternom, revisor.ternro FROM evacab inner join evadetevldor   on evadetevldor.evacabnro   = evacab.evacabnro"
			l_sql = l_sql & "	     and evadetevldor.evatevnro="& cevaluador & " inner join empleado revisor on revisor.ternro= evadetevldor.evaluador WHERE evacab.empleado = " & l_ternro & " and evacab.evaevenro="& l_evaevenro
			l_rs.Maxrecords = 1
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
				l_revisor = l_rs("empleg")
				l_rternro = l_rs("ternro")
				l_rterape = l_rs("terape")
				l_rternom = l_rs("ternom")
			end if 
			l_rs.Close
		else
			l_sql = "SELECT empleg, terape, ternom FROM empleado WHERE ternro=" & l_rternro
			l_rs.Maxrecords = 1
			rsOpen l_rs, cn, l_sql, 0
			if not l_rs.eof then
				l_rterape = l_rs("terape")
				l_rternom = l_rs("ternom")
				l_revisor = l_rs("empleg")
			end if	
			l_rs.Close
		end if
	end if ' de si es deloitte
	
	if l_rternro<> "" then
		'buscar el puesto REVISOR -------------------------------------------------------------
		l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro WHERE his_estructura.ternro=" & l_rternro
		if ccodelco=-1 then
		l_sql = l_sql & " AND   his_estructura.tenro = 46 "
		else
		l_sql = l_sql & " AND   his_estructura.tenro = 4 "
		end if
		l_sql = l_sql & " AND   his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
		l_rs.Maxrecords = 1
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then	
			l_rpuesto   = l_rs("estrdabr")
		else
			l_rpuesto = "--"
		end if	
		l_rs.Close
		
	end if
	
	'buscar la gerencia -----------------------------------------------------------------
	l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro WHERE his_estructura.ternro=" & l_ternro
	l_sql = l_sql & " AND   his_estructura.tenro = 6  AND   his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_gerencia = l_rs("estrdabr")
	else
		l_gerencia = "--"
	end if	
	l_rs.Close

	'buscar el puesto -----------------------------------------------------------------
	l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
	l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
	if ccodelco=-1 then
	l_sql = l_sql & " AND   his_estructura.tenro = 46 "
	else
	l_sql = l_sql & " AND   his_estructura.tenro = 4 "
	end if
	l_sql = l_sql & " AND   his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_puesto = l_rs("estrdabr")
		l_htetdesde= l_rs("htetdesde")
	else
		l_puesto = "--"
		l_htetdesde=""
	end if	
	l_rs.Close

	'Calcular antig en el puesto -----------------------------------------------------------------
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

	'buscar Sector -----------------------------------------------------------------
	l_sql = "SELECT estrdabr, htetdesde FROM his_estructura INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro WHERE his_estructura.ternro=" & l_ternro
	l_sql = l_sql & " AND   his_estructura.tenro = 2 AND his_estructura.htethasta IS NULL ORDER BY his_estructura.htetdesde DESC " 
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_sector = l_rs("estrdabr")
	else
		l_sector = "--"
	end if	
	l_rs.Close
	
	l_sql = "SELECT nrodoc FROM ter_doc WHERE ternro =" & l_ternro& " AND tidnro = 21 "  
	l_rs.Maxrecords = 1
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then	
		l_RUT = l_rs("nrodoc")
	else
		l_RUT = "--"
	end if	
 	l_rs.Close
	
	if cint(cdeloitte) =-1 then %> 
			<tr>
			<td colspan="2">
				<table  border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td nowrap><b>Empleado:</b>&nbsp;</td>
					<td nowrap>
			    		<input class="rev"  type="text" readonly value="<%= l_empleg %>" size="8" name="empleg">
						<input class="rev"  style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">
					</td>
					<%if l_evaproynro = "" then %>
					<td nowrap width="12%">&nbsp;
					<td nowrap>&nbsp;</td>
					<% else %>
					<td nowrap width="12%">	<b> Cliente:</b>&nbsp;</b>
					<td nowrap>
						<input class="rev" type="text" readonly value="<%= l_evaclicodext %>" size="8" name="empleg">
						<input class="rev" style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaclinom%>">
					</td>
					<% end if %>
				</tr>
				<tr>
					<td nowrap align="left"><b>Revisor:</b></td>
					<td colspan="1" nowrap>
					<input class="rev" readonly type="text"  value="<%= l_revisor %>" size="8" name="rempleg" readonly> 
					<input class="rev" readonly style="background : #e0e0de;" readonly type="text" name="revisor" size="35" maxlength="35" value="<%= l_rterape & ", " &l_rternom%>">
					</td>
					<% if l_evaproynro = "" then %>
					<td nowrap width="12%">&nbsp;
					<td nowrap>&nbsp;</td>
					<% else %>
					<td nowrap>	<b> Engagement:</b> &nbsp;</td>
					<td nowrap>
						<input class="rev" type="text" readonly value="<%= l_evaengcodext %>" size="8" name="empleg">
						<input class="rev" style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_evaengdesabr %>">
					</td>
					<% end if%>
				</tr>
			
				</table> 
			</td>
		</tr>	
		
	<% 
			l_linea = l_linea + 3
	   else
		if cint(ccodelco)=-1 then%>
		<tr>
			<td width="65%">
				<table  border="0" cellpadding="0" cellspacing="0" height="100%">
				<tr>
				    <td align="left"><b>Supervisado:</b></td>
					<td >
						<input class="rev"  type="text" readonly value="<%= l_empleg %>" size="8" name="empleg">
						<input class="rev"  style="background : #e0e0de;" readonly type="text" name="empleado" size="35" maxlength="35" value="<%= l_terape & ", " &l_ternom%>">
					</td>
				</tr>
				<tr>
					<td colspan="2"><b>Gerencia:</b>&nbsp;
					<input class="blanc" type="Text" name="gerencia" readonly size="30" value="<%=l_gerencia%>">
				    &nbsp;<b>Cargo:</b><input readonly class="blanc" type="Text" name="puesto" size="30" value="<%=l_puesto%>">
				    &nbsp;<b>RUT:</b>&nbsp;<input readonly  class="blanc" type="Text" name="rut" size="9" value="<%=l_rut%>">
				    </td>
				</tr>
				<tr>
					<td colspan="2">
						<b>Antig. en la Empresa:</b>&nbsp;
						<%dim l_texto, l_dia, l_mes, l_anio, l_hab
						l_texto = "--"
						if trim(l_ternro) <> "" then
							call antigfec (l_ternro, date , l_dia, l_mes, l_anio, l_hab)
							if l_anio = "" or l_anio = 0 then
								if l_mes = "" or l_mes = 0 then
									l_texto = l_dia & " día/s."
								else
									l_texto = l_mes & " mes/es " & l_dia & " día/s."
								end if	
							else
								l_texto = l_anio & " año/s " & l_mes & " mes/es " 
							end if
						end if
						%>  
						<input class="blanc" readonly type="Text" name="antiguedad" size="25" value='<%=l_texto%>'>
						&nbsp;<b>en el Cargo:</b>&nbsp;
						<input class="blanc" readonly type="Text" name="antiguedadpuesto" value="<%=l_antiguedadpuesto%>" size="25">
					</td>
				</tr>
				</table> 
				</td>
				<td>
						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
						    <td align="right"><b>Per&iacute;odo Desde:</b></td>
							<td><input readonly class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperdesde %>"></td>
						</tr>	
						<tr>	
						    <td align="right" ><b>Hasta:</b></td>
							<td><input class="blanc" readonly type="Text" name="usr" size="10" value="<%= l_evaperhasta %>"></td>
						</tr>
						</table> 
				</td>
			</tr>
			<tr>
				<td colspan="3">
					<table  width="98%" border="0" cellpadding="0" cellspacing="0">
					<tr>
					    <td align="right"><b>Supervisor:</b></td>
						<td>
							<input class="rev" readonly type="text" value="<%= l_revisor %>" size="6" name="rempleg" readonly> 
							<input class="rev" style="background : #e0e0de;" readonly type="text" name="revisor" size="30" maxlength="30" value="<%= l_rterape & ", " &l_rternom%>">
							&nbsp;
							<b>Cargo:</b>&nbsp;<input readonly class="blanc" type="Text" name="evalpuesto" size="25" value="<%=l_rpuesto%>">
						</td>
						<td>
						</td>
					</tr>
					</table> 
				</td>
			</tr>
		<% else %> 
		<tr>
		<td  colspan="2">
		<table border="1" style=" width:100%;">
			<tr>
			<td>
				<table  border="0" cellpadding="0" cellspacing="0" style=" width:100%;">
				<tr>
				    <td align="left" colspan="3"><b>Empleado:</b>&nbsp;
						<input type="text" style="background : #e0e0de;" readonly value="<%= l_empleg %>" size="4" name="empleg">
						<input style="background : #e0e0de;" readonly type="text" name="empleado" size="32" maxlength="32" value="<%= l_terape & ", " &l_ternom%>">
					</td>
				</tr>
				<tr>			
					<td align="left">
					<b>Gerencia</b> <br>
					<input class="blanc" readonly type="Text" name="usr" value="<%=l_gerencia%>" size="35">
					</td>
					<td align="left">
					<b><%if cejemplo=-1 then%>&nbsp;<%else%>Puesto<%end if%></b><br>
					<%if cejemplo=-1 then%>&nbsp;<%else%><input readonly class="blanc" type="Text" name="usr" value="<%=l_puesto%>" size="30"><%end if%>
					</td>
					<td align="left" ><b>Sector</b><br>
					<input class="blanc" readonly  type="Text" name="usr" value="<%=l_sector%>" size="29">
					</td>

					</tr>
				</table> 
				
			</td>
			<td>
				<table  border="0" cellpadding="0" cellspacing="0" style=" width:100%;">
				<tr>
				    <td align="right"><b>Per&iacute;odo Desde:</b></td>
					<td><input class="blanc" readonly  type="Text" name="usr" size="10" value="<%= l_evaperdesde %>"></td>
				</tr>	
				<tr>	
				    <td align="right" ><b>Hasta:</b></td>
					<td><input readonly class="blanc" type="Text" name="usr" size="10" value="<%= l_evaperhasta %>"></td>
				</tr>
				</table> 
			
			</td>
			</tr>
		
		<tr>
			<td colspan="2">
				<table border="0" cellpadding="0" cellspacing="0" style=" width:100%;">
				<tr>
				    <td align="left">
					<b><%if cejemplo=-1 then%>Evaluador:<%else%>Revisor:<%end if%></b>&nbsp;
					<%if cejemplo<>-1 then%>
					<input readonly class="rev" type="text" style="background : #e0e0de;" readonly value="<%= l_revisor %>" size="8" name="rempleg"> 
					<%end if%>
					<input class="rev" style="background : #e0e0de;" readonly type="text" name="revisor" size="35" maxlength="35" value="<%= l_rterape & ", " &l_rternom%>">
					</td>
				    <td align="center">
						<b><%if cejemplo=-1 then%>&nbsp;<%else%>Puesto:<%end if%></b>&nbsp;<%if cejemplo=-1 then%>&nbsp;<%else%><input readonly class="blanc" type="Text" name="evalpuesto" value="<%=l_rpuesto%>" size="30"><%end if%>
					</td>
				</tr>
					</table> 
			
			</td></tr>	
		 
		 
		 </table>
		</td>
		</tr>
		<%
			l_linea = l_linea + 6
		end if ' si no es cdelolitte ni codelco- Estandar
	   end if

			
	
	Dim	l_evaseccnro 
	Dim	l_titulo 
	Dim l_tipsecprogres
	
	l_sql = "SELECT evasecc.titulo, evasecc.orden, evasecc.evaseccnro, evatiposecc.tipsecprogres  "
	l_sql = l_sql & " FROM evadet INNER JOIN evasecc ON evadet.evaseccnro=evasecc.evaseccnro INNER JOIN evatiposecc ON evasecc.tipsecnro=evatiposecc.tipsecnro "
	if l_tieneobj=0 then
	l_sql = l_sql & " AND evatiposecc.tipsecobj=0"
	end if
	l_sql = l_sql & " WHERE evadet.evacabnro =" & l_evacabnro & " ORDER BY orden "
	rsOpen l_rssecc, cn, l_sql, 0 
	
	do until l_rssecc.eof
			l_evaseccnro	= l_rssecc("evaseccnro")
			l_titulo		= l_rssecc("titulo")
			l_tipsecprogres	= l_rssecc("tipsecprogres")
			if trim(l_tipsecprogres)<>"" then%>
			<tr>
				<td colspan="2"><b><br><br>Secci&oacute;n&nbsp;<%= l_titulo %></b></td>
			</tr>
			<tr>
				<td colspan="2">
				<%
					l_linea = l_linea + 4
			end if
			'response.write ("<script>alert('"&l_tipsecprogres&"')</script>")
			select case trim(l_tipsecprogres)
				case "resumen_planaccion_eva_00.inc"
					resumen_plan_accion
				case "resumen_plansmart_eva_00.inc"
					resumen_plansmart
				case "resumen_plansmart_eva_ABN.inc"
					resumen_plansmartABN
				case "resumen_cardinales_eva_00.inc"
					resumen_cardinales
				case "resumen_objetivos_eva_00.inc"
					resumen_objetivos
				case "resumen_objetivos_plan_eva_00.inc"
					resumen_objetivos_plan				
				case "resumen_general_eva_ABN.inc"
					resumen_generalABN
				case "resumen_objetivossmart_eva_ABN.inc"
					resumen_objetivossmartABN
				case "resumen_objetivossmart_eva_00.inc"  ' *********
					resumen_objetivossmart	
				case "resumen_objetivossmart_eva.inc"
					resumen_objetivossmart				
				case "resumen_notas_eva_00.inc"
					resumen_notas
				case "resumen_vistos_eva_00.inc"
					resumen_vistos
				case "resumen_totales_eva_00.inc"
					resumen_totales
				case "resumen_resultados_eva_ABN.inc"
					resumen_resultadosABN
				case "resumen_resultados_eva_00.inc"
					resumen_resultados
				case "resumen_grafico_eva_00.inc"
					resumen_grafico
				'Para Deloitte	
				case "resumen_datosadm_eva_00.inc"
					resumen_datosadm
				case "resumen_resultadosyarea_eva_00.inc"
					resumen_resultadosyarea
				case "resumen_calificobj_eva_00.inc"
					resumen_calificobj
				' case "resumen_gralobj_eva_00.inc"	 'resumen_gralobj 
				case "resumen_areacom_eva_00.inc"
					resumen_areacom
				case "resumen_objcom_eva_00.inc"
					resumen_objcom
				case "resumen_areacomRDP_eva_00.inc"
					resumen_areacomRDP
				case "resumen_datosadmRDP_eva_00.inc"
					resumen_datosadmRDP
				case "resumen_calificobjRDP_eva_00.inc"
					resumen_calificobjRDP
				case "resumen_calificcompRDP_eva_00.inc"
					resumen_calificcompRDP
				case "resumen_calificgralRDP_eva_00.inc"
					resumen_calificgralRDP
				'Para Codelco
				case "resumen_borrador_eva_COD.inc"
					resumen_borrador
				case "resumen_compromisos_eva_COD.inc"
					resumen_compromisos
				case "resumen_cierre_eva_COD.inc"
					resumen_cierre
				case "resumen_cierreEva_eva_COD.inc"
					resumen_cierreEva
				case "resumen_actividades_eva_COD.inc"
					resumen_actividades
				case "resumen_retroalimentacion_eva_COD.inc"
					resumen_retroalimentacion
				case "resumen_evalborrador_eva_COD.inc"
					resumen_evalborrador
				case "resumen_evaluacion_eva_COD.inc"
					resumen_evaluacion	
				case "resumen_compxestr_eva_00.inc"
					resumen_compxestr				
				case "resumen_calificcompSI_eva_00.inc"
					resumen_calificcompSI
				case "resumen_calificcompSE_eva_00.inc"
					resumen_calificcompSE	
			end select

			if trim(l_tipsecprogres)<>"" then%>
				</td>
			</tr>
			<%
			end if
			l_rssecc.MoveNext
			l_tipsecprogres=""
	loop
	l_rssecc.Close
	
	if l_i <> UBound(l_arrternro)  then
%>	
	</table>
	<!-- <tr> <td colspan=2> -->
		<p style='page-break-before:always'></p>
	<!--<br></td>	</tr> -->
	<table border="1" cellpadding="0" cellspacing="0" style=" width: 100%;">
<%
	end if

Next 

set l_rs=nothing
set l_rssecc=nothing
cn.Close
set cn = Nothing

%>
</table>
</form>
</body>
</html>
