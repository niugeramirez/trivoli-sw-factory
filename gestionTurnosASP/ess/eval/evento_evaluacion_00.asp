<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Modificado: 12-10-2005 - Leticia A. - Adaptarlo a Autogestion.
Modificado: 31-10-2005 - CCR. - Declarar variable l_empleg sin declarar...
-->
<% 
' Variables
dim l_empleg
' de parametros entrada
  
' de base de datos  
  Dim l_sql
  Dim l_rs

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "C&oacute;digo:;Descripci&oacute;n:;Tipo Evento:;Formulario Evaluaci&oacute;n:;Fecha Evaluaci&oacute;n:;Desde:;Hasta:"
  l_Campos    = "v_evaevento.evaevenro;v_evaevento.evaevedesabr;evatipoevento.evatipevedabr;evatipoeva.evatipdesabr;v_evaevento.evaevefecha;v_evaevento.evaevefdesde;v_evaevento.evaevefhasta"
  l_Tipos     = "N;T;T;T;F;F;F"

' Orden
  l_Orden     = "C&oacute;digo:;Descripci&oacute;n:;Tipo Evento:;Formulario Evaluaci&oacute;n:;Fecha Evaluacion:;Desde:;Hasta:"
  l_CamposOr  = "v_evaevento.evaevenro;v_evaevento.evaevedesabr;evatipoevento.evatipevedabr;evatipoeva.evatipdesabr;v_evaevento.evaevefecha;v_evaevento.evaevefdesde;v_evaevento.evaevefhasta"

' parametros de entrada---------------------------------------
l_empleg = Request.QueryString("empleg")

%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Eventos <%if ccodelco=-1 then%>del Ciclo de Gestión del Desempeño<%else%>Evaluaci&oacute;n<%end if%> - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function orden(pag)
{
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{	// filtro_browse.asp?
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}	


function param(){
	return ('empleg=<%= l_empleg%>');
}

function pantalla(){
	document.datos.pantalla.value=screen.availWidth;
}

//window.resizeTo(700,350); 	   
</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" height="100%" onload="pantalla();">
<form name=datos>
	<input type=hidden name=seleccion>
	<input type=hidden name=pantalla>
</form>

<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<th align="left" width="30%">Eventos <%if ccodelco=-1 then%>del Ciclo de Gestión del Desempeño<%else%>Evaluaci&oacute;n<%end if%></th>
		<th align="right">
			<a class=sidebtnABM href="Javascript:abrirVentana('evento_evaluacion_02.asp?Tipo=A','',505,450)">Alta</a> &nbsp;
			<a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'evento_evaluacion_04.asp?evaevenro=' + document.ifrm.datos.cabnro.value)">Baja</a>&nbsp;
			<%if ccodelco=-1 then%>
			<a class=sidebtnABM href="Javascript:abrirVentanaVerif('evento_evaluacion_02.asp?Tipo=M&evaevenro=' + document.ifrm.datos.cabnro.value,'',505,400)">Modifica</a>&nbsp;
			<%else%>
			<a class=sidebtnABM href="Javascript:abrirVentanaVerif('evento_evaluacion_02.asp?Tipo=M&evaevenro=' + document.ifrm.datos.cabnro.value,'',505,450)">Modifica</a>&nbsp;
			<%end if%>
			<a class=sidebtnSHW href="Javascript:orden('/ess/ess/eval/evento_evaluacion_01.asp');">Orden</a>&nbsp;
			<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/eval/evento_evaluacion_01.asp');">Filtro</a>&nbsp;
		</th>
	</tr>
	<tr style="border-color :CadetBlue;">
		<!--<td align="left" width="30%" class="barra"></td>-->
		<td align="right" class="th2" colspan="2">
			<%if ccodelco=-1 then%>
			<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('eventos_porcentajeobj_eva_00.asp?evaevenro=' + document.ifrm.datos.cabnro.value,'',550,380)">Ponderación Tipos Compromisos</a>&nbsp;
			<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('form_carga_eva_COD_00.asp?evaevenro=' + document.ifrm.datos.cabnro.value+'&pantalla=' + document.datos.pantalla.value,'',800,580);">Gestionar el Desempeño</a>&nbsp;
			<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('relacionar_empleados_eva_00.asp?evaevenro='+document.ifrm.datos.cabnro.value+'&obj=document.datos.seleccion','',600,600);">Relacionar Supervisados</a>&nbsp;
			<%else%>
				<% if cdeloitte=-1 then %>
				<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('form_carga_eva_DEL_00.asp?evaevenro=' + document.ifrm.datos.cabnro.value+'&pantalla=' + document.datos.pantalla.value,'',800,580);">Evaluar</a> &nbsp;
				<% else %>
				<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('form_carga_eva_00.asp?evaevenro=' + document.ifrm.datos.cabnro.value+'&pantalla=' + document.datos.pantalla.value,'',800,580);">Evaluar</a>&nbsp;
				<% end if%>
				<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('relacionar_empleados_eva_00.asp?evaevenro='+document.ifrm.datos.cabnro.value+'&obj=document.datos.seleccion','',600,400);">Relacionar Empleados</a> &nbsp;
			<%end if%>
			<%if cejemplo<>-1 and ccodelco<>-1 then%>
			<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('etapa_form_masiva_eva_00.asp?evaevenro=' + document.ifrm.datos.cabnro.value,'',460,425)">Cambio Etapa</a> &nbsp;
			<%end if%>
			<a class=sidebtnSHW href="Javascript:abrirVentanaVerif('monitor_evento_eva_00.asp?evaevenro='+document.ifrm.datos.cabnro.value,'',800,400);">Monitor del Evento</a>&nbsp;
		</td>
	</tr>
	<tr valign="top" height="85%">
		<td colspan="2" style="">
			<iframe name="ifrm" src="evento_evaluacion_01.asp?empleg=<%=l_empleg%>" width="100%" height="100%"></iframe> 
		</td>
	</tr>
	<tr>
		<td colspan="2" height="2%"></td>
	</tr>
</table>

</body>
</html>
