<% Option Explicit %>
<%	'<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->  %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : licencias_emp_gti_00.asp
Descripcion    : Modulo que se encarga administrar las licencias de un empleado
Creacion       : 24/03/2004
Autor          : Scarpa D.
Modificacion   :
  06/05/2004 - Scarpa D. - Cambio en el tamaño de las ventanas de alta/modificacion
  18/10/2004 - Scarpa D. - Cambio en el tamaño de las ventanas de alta/modificacion
  11/10/2005 - Leticia Amadio - Cambio de las opciones de filtro y orden para Autogestion.
  27-11-2006 - Diego Rosso - Cambio para que no se puedan modificar ni eliminar vacaciones que ya tengan pago/desc
  09/03/2007 - Lisandro Moro - Se Agrego el boton Salida excel.
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

' Variables
Dim l_ternro
dim l_rs
dim l_sql
Dim l_empleg
Dim l_habilitar_estado

l_habilitar_estado = (Session("empleg") <> l_ess_empleg)

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_Campos    = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
  l_Tipos     = "T;T;F;F"

' Orden
  l_Orden     = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_CamposOr  = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
  
  l_empleg = request("empleg")
%>
<script>
function orden(pag)
{
	// /serviciolocal/ess/shared/asp/..
  abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)  
}

function filtro(pag)
{
  abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
} 

function param(){
	return ('empleg=<%=l_empleg%>');
}

function Baja(){
	
	if (document.ifrm.datos.ModEli.value != "0") // Diego Rosso
	{
		alert("La licencia tiene un pago/desc asociada. No puede ser eliminada");
		return;
	}
    <%if not l_habilitar_estado then%>
	if (document.ifrm.datos.estado.value != "1")
		alert("La licencia no puede ser eliminada");
	else
	<%end if%>
		eliminarRegistro(document.ifrm,'licencias_emp_gti_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);
}

function Modificacion(){
	
	if (document.ifrm.datos.ModEli.value != "0") // Diego Rosso
	{
		alert("La licencia tiene un pago/desc asociada. No puede ser Modificada");
		return;
	}

    <%if not l_habilitar_estado then%>
	if (document.ifrm.datos.estado.value != "1")
		alert("La licencia no puede ser Modificada");
	else
	<%end if%>
		abrirVentanaVerif('licencias_emp_gti_02.asp?tipo=M&empleg=<%= l_empleg%>&cabnro=' + document.ifrm.datos.cabnro.value,'',570,300,'resizable=no');
}

function alta(){
    abrirVentana('licencias_emp_gti_02.asp?tipo=A&empleg=<%= l_empleg%>','',570,300,'resizable=no')
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("licencias_emp_gti_excel.asp?empleg=<%= l_empleg%>&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}

</script>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Licencias - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="95%">
        <tr>
          <th colspan="2" align="left">Licencias</th>
          <th colspan="2" style="text-align: right">
		  <a class=sidebtnABM href="Javascript:alta();">Alta</a>
		  <a class=sidebtnABM href="Javascript:Baja();">Baja</a>
		  <a class=sidebtnABM href="Javascript:Modificacion();">Modifica</a>
		  &nbsp;&nbsp;&nbsp;
		  <a class=sidebtnSHW href="Javascript:llamadaexcel();">Excel</a>
		  &nbsp;&nbsp;&nbsp;
		  <!-- <a class=sidebtnSHW href="Javascript:orden('/serviciolocal/licencias_emp_gti_01.asp');">Orden</a> -->
          <a class=sidebtnSHW href="Javascript:orden('/ess/ess/gti/licencias_emp_gti_01.asp');">Orden</a>
	      <a class=sidebtnSHW href="Javascript:filtro('/ess/ess/gti/licencias_emp_gti_01.asp');">Filtro</a>
		  &nbsp;&nbsp;&nbsp;
		  </th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="licencias_emp_gti_01.asp?empleg=<%= l_empleg%>" width="100%" height="100%"></iframe> 		  
	      </td>
        </tr>
        <tr valign="top">
          <td colspan="4"  height="20">
	      </td>
        </tr>

      </table>
</body>
</html>
