<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<% 
'=========================================================================================
'Archivo  : historicas_evldr_eva_ag_00.asp
'Objetivo : Listados de formularios de evaluacion - cerrados o aprobados - rol:evaluador
'Fecha	  : 12-05-2006
'Autor	  : Leticia Amadio
'Modificación: 04-10-2006 - Leticia Amadio - Adecuarlo a autogestion
'========================================================================================
on error goto 0

' Variables
' de parametros entrada
  Dim l_empleg   ' es el EMPLEG viene de autogestion
  Dim l_tipoeval ' si es deloitte viene con datos
  
' de uso local  
  dim l_nombre
  dim l_ternro
	  
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
  l_etiquetas = "C&oacute;digo:;Evento:;Fecha Aprobaci&oacute;n:;Hora Aprobaci&oacute;n:"
  l_Campos    = "evaevento.evaevenro;evaevento.evaevedesabr;evacab.fechaapro;evacab.horaapro"
  l_Tipos     = "N;T;F;T"

' Orden
  l_Orden     = "C&oacute;digo:;Evento:;Fecha Aprobaci&oacute;n:;Hora Aprobaci&oacute;n:" 
  l_CamposOr  = "evaevento.evaevenro;evaevento.evaevedesabr;evacab.fechaapro;evacab.horaapro"

' parametros de entrada---------------------------------------
l_empleg = l_ess_empleg
l_tipoeval = Request.QueryString("tipoeval")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro, terape, ternom, terape2, ternom2 FROM empleado "
l_sql = l_sql & " WHERE empleg = " & l_empleg
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_ternro= l_rs("ternro")
	l_nombre = l_rs("terape")
	if trim(l_rs("terape2"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("terape2"))
	end if	
	if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & ","
	end if	
	if trim(l_rs("ternom"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom"))
	end if	
	if trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom2"))
	end if	
end if
l_rs.close
set l_rs=nothing
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%if ccodelco=-1 then%>Procesos de Gesti&oacute;n<%else%>Formularios de Evaluaci&oacute;n- Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;<%end if%></title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
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

/*function param(){ Deloitte
	chequear=""
	<%'if l_tipoeval="RDP" or l_tipoeval="SERVI" then %>
		chequear= "tipoeval=<%'=l_tipoeval%>&"
	<%'end if %>
	chequear= chequear + "ternro=<%'= l_ternro %>";
	return chequear;
}*/

function param(){
	chequear= "empleg=<%= request.querystring("empleg") %>";
	return chequear;
}

function pantalla(){
	document.datos.pantalla.value=screen.availWidth;
}

</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="pantalla();">
<form name=datos>
	<input type=hidden name=pantalla>
</form>
<table border="0" cellpadding="0" cellspacing="0" height="100%">
<tr style="border-color :CadetBlue;" height="5%">
	<th colspan="2" align="left" class="th2"><%if ccodelco=-1 then%>Procesos <%else%>Formularios <%end if%>Hist&oacute;ricos de <%=l_nombre%> (rol: Evaluador)</th>
</tr>
<tr height="5%">
	<td colspan="2" align="right" class="th2">
	<% if cint(cdeloitte) =-1 then %>
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_DEL_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&llamadora=consulta','',800,600)">Ver Evaluaci&oacute;n</a>
	<% else %>
		<% if ccodelco=-1 then%>
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_COD_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&llamadora=consulta','',800,600)">Ver Evaluaci&oacute;n</a>
		<% else%>
		<a class=sidebtnABM href="Javascript:abrirVentanaVerif('form_carga_eva_ag_00.asp?evacabnro=' + document.ifrm.datos.cabnro.value+'&evaevenro='+ document.ifrm.datos.evaevenro.value+'&empleg='+ document.ifrm.datos.empleg.value+'&pantalla=' + document.datos.pantalla.value+'&llamadora=consulta','',800,600)">Ver Evaluaci&oacute;n</a>
		<% end if%>
	<% end if%>
	&nbsp;
	<a class=sidebtnSHW href="Javascript:orden('/ess/ess/eval/historicas_evldor_eva_ag_01.asp');">Orden</a>
	<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/eval/historicas_evldor_eva_ag_01.asp');">Filtro</a>
	&nbsp;
	</td>
</tr>

<tr valign="top" height="95%">
   <td colspan="3" style="">
   <iframe name="ifrm" src="historicas_evldor_eva_ag_01.asp?empleg=<%'=request.querystring("empleg")%>" width="100%" height="100%"></iframe> 
  </td>
</tr>

</table>
</body>
</html>
