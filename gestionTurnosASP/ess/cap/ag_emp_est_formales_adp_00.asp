<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'Archivo		: emp_est_formales_adp_00.asp
'Descripción	: relacion empleado estudios
'Autor			: lisandro moro
'Fecha			: 09/08/2003 
'Modificado:
'	Alvaro Bayon - 15-09-2003 - Llamado a eliminar
'								Se pasa como parámetro el nivel						
'Modificado 19-09-2003 - CCROssi   - aplicar funcion MostrarBoton 
'Modificado 22-09-2003 - CCROssi   - agregar Filtro Orden y Excel
'Modificado 25-02-2004 - Scarpa D. - Se agrego la opcion de estudio actual
'			05-03-2004 - Scarpa D. - Cambio del tamaño de la ventana
'			03/11/2006 - Lisandro Moro - Se comento la fn MostrarBoton por un error en WalMart. 
' ----------------------------------------------------------------------------------
on error goto 0

dim l_ternro
dim l_empleg
dim l_empleado


dim l_rs
dim l_sql
dim l_sql2
dim NombrePuesto
dim l_puenro

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "Nivel:;Titulo:;Institucion:;Carrera;Fecha Desde:;Fecha Hasta:;"
  l_Campos    = "nivdesc;titdesabr;instdes;carredudesabr;capfecdes;capfechas"
  l_Tipos     = "T;T;T;T;F;F"

' Orden
  l_Orden     = "Nivel:;Titulo:;Institucion:;Carrera;desde:;Hasta:;"
  l_CamposOr  = "nivdesc;titdesabr;instdes;carredudesabr;capfecdes;capfechas"

l_ternro = l_ess_ternro


Set l_rs = Server.CreateObject("ADODB.RecordSet")
  
l_sql = " SELECT estructura.estrdabr, puesto.puenro "
l_sql = l_sql & " FROM his_estructura "
l_sql = l_sql & " INNER JOIN puesto ON puesto.estrnro = his_estructura.estrnro "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro=4 AND his_estructura.ternro=" & l_ternro & " AND his_estructura.htetdesde <= " & cambiafecha(date(),"YMD",true) & " AND ((" & cambiafecha(date(),"YMD",true) & " <= his_estructura.htethasta) OR (his_estructura.htethasta IS NULL)) "

rsOpen l_rs, cn, l_sql, 0

If Not l_rs.EOF Then
	NombrePuesto = l_rs("estrdabr")
	l_puenro     = l_rs("puenro")
else 	
	NombrePuesto = "Sin Puesto"
	l_puenro     = 0
end If

l_rs.close
  
  
%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Estudios Formales del Empleado&nbsp;-&nbsp;Administraci&oacute;n de Personal - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>
function Aceptar(){
	window.close();
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("ag_emp_est_formales_adp_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value) + "&ternro=<%=l_ternro%>&estado=" + document.all.estado.value,'execl',250,150);
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}
 	   
function param(){
	chequear= "ternro=<%= l_ternro %>&estado=" + document.all.estado.value;
	return chequear;
}

function cambioEstado(){
  document.ifrm.location = "ag_emp_est_formales_adp_01.asp?ternro=<%= l_ternro %>&estado=" + document.all.estado.value + "&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value);
}

</script>

</head>
<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr>
	<th align="left" nowrap>Estudios Formales</th>
	<th nowrap align="right">
	      <b>Estado:</b>
		  <select name="estado" onChange="javascript:cambioEstado();">
		    <option value="">Todos
		    <option value="-1">Actuales
		    <option value="0">Futuros
		  </select>
		  &nbsp;&nbsp;
		  <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentana('ag_emp_est_formales_adp_02.asp?tipo=A&ternro="&l_ternro&"','',700,350)","Alta") %>
		  <% 'call MostrarBoton ("sidebtnABM", "Javascript:eliminarRegistro(document.ifrm,'ag_emp_est_formales_adp_04.asp?ternro="&l_ternro&"&nivnro='+document.ifrm.datos.nivnro.value+'&titnro='+document.ifrm.datos.titnro.value+'&instnro='+document.ifrm.datos.instnro.value+'&carredunro='+document.ifrm.datos.carredunro.value)","Baja") %>
		  <% 'call MostrarBoton ("sidebtnABM", "Javascript:abrirVentanaVerif('ag_emp_est_formales_adp_02.asp?tipo=M&ternro="&l_ternro&"&nivnro='+document.ifrm.datos.nivnro.value+'&titnro='+document.ifrm.datos.titnro.value+'&instnro='+document.ifrm.datos.instnro.value+'&carredunro='+document.ifrm.datos.carredunro.value,'',700,350)","Modifica") %>
		  <a class=sidebtnABM href="Javascript:abrirVentana('ag_emp_est_formales_adp_02.asp?tipo=A&ternro=<%= l_ternro %>','',700,350);">Alta</a>
		  <a class=sidebtnABM href="Javascript:eliminarRegistro(document.ifrm,'ag_emp_est_formales_adp_04.asp?ternro=<%= l_ternro %>&nivnro='+document.ifrm.datos.nivnro.value+'&titnro='+document.ifrm.datos.titnro.value+'&instnro='+document.ifrm.datos.instnro.value+'&carredunro='+document.ifrm.datos.carredunro.value);">Baja</a>
		  <a class=sidebtnABM href="Javascript:abrirVentanaVerif('ag_emp_est_formales_adp_02.asp?tipo=M&ternro=<%= l_ternro %>&nivnro='+document.ifrm.datos.nivnro.value+'&titnro='+document.ifrm.datos.titnro.value+'&instnro='+document.ifrm.datos.instnro.value+'&carredunro='+document.ifrm.datos.carredunro.value,'',700,350);">Modifica</a>
		  &nbsp;
		  <a class=sidebtnSHW href="Javascript:orden('../../cap/ag_emp_est_formales_adp_01.asp');">Orden</a>
		  <a class=sidebtnSHW href="Javascript:filtro('../../cap/ag_emp_est_formales_adp_01.asp')">Filtro</a>
		  &nbsp;
  		  <a class=sidebtnSHW href="Javascript:llamadaexcel();">Excel</a>
	</th>
</tr>
<tr>
	<td colspan="2">
		<table  border="0" cellpadding="0" cellspacing="0" >
			<tr>
			    <td nowrap align="right" ><b>Puesto:</b></td>
				<td><input style="background : #e0e0de;" type="text" name="convenio" size="71" maxlength="50" value='<%= NombrePuesto %>' readonly></td>
			    <td nowrap align="right">&nbsp;</td>
				<td >&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" height="90%">
	<iframe name="ifrm" height="100%" src="ag_emp_est_formales_adp_01.asp?ternro=<%= l_ternro %>" width="100%"></iframe> 
    </td>
</tr>
</table>
</html>
<%
Set l_rs = nothing
cn.close
set cn = nothing
%>
