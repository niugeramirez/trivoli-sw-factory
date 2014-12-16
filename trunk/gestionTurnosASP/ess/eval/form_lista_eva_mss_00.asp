<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'=====================================================================================
'Archivo  : form_lista_eva_mss_00.asp
'Objetivo : Listados de formularios de evaluacion - autogestion para MSS 
'Fecha	  : 11-11-2005
'Autor	  : Leticia Amadio
'Modificado:
'=====================================================================================

' Variables
' de parametros entrada
  Dim l_empleg
  Dim l_logeadoempleg   ' es el EMPLEG viene de autogestion - o del MSS
  dim l_esgerente
  
' de uso local  
  Dim l_logeadoternro   ' es el TERNRO del empleg que viene de autogestion - o del MSS
  dim l_evatevnro  
  
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
  if ccodelco=-1 then
  l_etiquetas = "Per&iacute;odo:;Formulario:;Supervisado a Evaluar:"
  else
  l_etiquetas = "Evento:;Formulario:;Empleado a Evaluar:"
  end if
  l_Campos    = "evaevento.evaevedesabr;evatipoeva.evatipdesabr;v_empleado.terape"
  l_Tipos     = "T;T;T"

' Orden
  if ccodelco=-1 then 
  l_Orden      = "Per&iacute;odo:;Formulario:;Supervisado a Evaluar:"
  else 
  l_Orden     = "Evento:;Formulario:;Empleado a Evaluar:"
  end if 
  
  l_CamposOr  = "evaevento.evaevedesabr;evatipoeva.evatipdesabr;v_empleado.terape"


l_empleg = request("empleg") 
l_logeadoempleg = session("empleg") 'l_ess_empleg

' __________________________________________________
'l_logeadoempleg = Session("empleg")				
'if trim(l_logeadoempleg)="" then					
	'l_logeadoempleg = Request.QueryString("empleg")
	'Session("empleg")=l_logeadoempleg				
'end if												
' __________________________________________________

' ___ se fija si es gerente, ie si es autoevaluador o evaluador __________
Set l_rs = Server.CreateObject("ADODB.RecordSet")						  
l_sql = "SELECT ternro FROM v_empleado WHERE  empleg = " & l_logeadoempleg
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_logeadoternro = l_rs("ternro")
end if
l_rs.Close

l_evatevnro=0	
l_sql = "SELECT evatevnro FROM evadetevldor WHERE  evaluador=" & l_logeadoternro
l_sql = l_sql & " AND evatevnro =" & cevaluador 
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
   l_evatevnro = l_rs("evatevnro")
end if
l_rs.Close
set l_rs=nothing

l_esgerente=-1
if cint(l_evatevnro)=cint(cautoevaluador) or cint(l_evatevnro)=cint(cevaluador) then
	l_esgerente=0
end if	
' ____________________________________________________________________________
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Proceso de Gesti&oacute;n de Desempe&ntilde;o  -  Gesti&oacute;n de Desempe&ntilde;o - <%if ccodelco<>-1 then%>RHPro &reg;<%end if%></title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function filtro(pag)
{
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

/* function param(){
	chequear= "logeadoternro=<%'= l_logeadoternro %>";
	return chequear;
} */

function param(){
	return ('empleg=<%= l_empleg%>');
}

function pantalla(){
	document.datos.pantalla.value=screen.availWidth;
}

<%if cint(ccodelco)=-1 then%>
function deshabbtn(){
	document.all.btn1.className="sidebtnDSB";
	document.all.btn2.className="sidebtnDSB";
	document.all.btn3.className="sidebtnDSB";
	document.all.btn1.disabled=true;
	document.all.btn2.disabled=true;
	document.all.btn3.disabled=true;
}

function habbtn(){
	document.all.btn1.className="sidebtnSHW";
	document.all.btn2.className="sidebtnSHW";
	document.all.btn3.className="sidebtnSHW";
	document.all.btn1.disabled=false;
	document.all.btn2.disabled=false;
	document.all.btn3.disabled=false;
}
<%end if%>
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="pantalla();<%if cint(ccodelco)=-1 then%>deshabbtn();<%end if%>">
<form name=datos>
	<input type=hidden name=pantalla>
</form>
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;" height="5%">
	<th align="left">Proceso de Gesti&oacute;n de Desempe&ntilde;o </th>
	<th align="right">
	<!--   SACAR POR AHORA..?? -->
	<%if cint(cejemplo)=-1 and cint(l_esgerente) = -1 then%>
		<a class=sidebtnABM href="Javascript:abrirVentana('tieneobjetivos_eva_ag_00.asp?empleg=<%=l_logeadoempleg%>&listainicial='+document.ifrm.datos.listainicial.value,'',500,300)">Definir Empleados con Objetivos</a>
	<%end if%>
	
	<!--  Comento lo customizado..
	<%'if cint(cdeloitte)=-1 then%>
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_DEL_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&logeadoempleg=<%=l_logeadoempleg%>' ,'',800,600)">Gestionar</a>
	<%'else
		'if cint(ccodelco)=-1 then%>
			<a class=sidebtnDSB name="btn1" href="Javascript:abrirVentanaVerif('objiniciales_eva_00.asp?empleg=<%=l_logeadoempleg%>&evaevenro='+ document.ifrm.datos.evaevenro.value,'',650,300)">Generar Compromisos Inciales</a>
			<a class=sidebtnDSB name="btn2" href="Javascript:abrirVentana('tieneobjetivos_eva_ag_00.asp?empleg=<%=l_logeadoempleg%>&listainicial='+document.ifrm.datos.listainicial.value,'',500,300)">Definir Supervisados con Compromisos Iniciales</a>
			<a class=sidebtnDSB name="btn3" href="Javascript:abrirVentana('rep_emp_formulario_eva_00.asp?llamadora=AUTO&logeadoternro=<%=l_logeadoternro%>&evaevenro='+ document.ifrm.datos.evaevenro.value,'',800,600)">Reporte Formulario</a>
			<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_COD_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&logeadoempleg=<%=l_logeadoempleg%>' ,'',800,600)">Gestionar</a>-->
		<%'else%>
			<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&logeadoempleg=<%=l_logeadoempleg%>' ,'',800,600)">Gestionar</a>
		<%'end if
	'end if%>
		<a class=sidebtnABM href="Javascript:abrirVentana('historicas_eva_ag_00.asp?empleg=<%= request.querystring("empleg")%>','',500,300)">Hist&oacute;ricas</a>
		&nbsp;
		<a class=sidebtnSHW href="Javascript:orden('/ess/ess/eval/form_lista_eva_mss_01.asp');">Orden</a>
		<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/eval/form_lista_eva_mss_01.asp');">Filtro</a>
		&nbsp;
	</th>
</tr>

<tr valign="top" height="95%">
   <td colspan="2" style="">
   <!-- <iframe name="ifrm" src="form_lista_eva_ag_01.asp?logeadoternro=<% '=l_logeadoternro%>" width="100%" height="100%"></iframe> -->
   <iframe name="ifrm" src="form_lista_eva_mss_01.asp?empleg=<%=l_empleg%>" width="100%" height="100%"></iframe> 
   </td>
</tr>

</table>
</body>
</html>
<%
cn.Close
set cn = Nothing
%>