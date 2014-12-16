<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'            14-10-2005 - Leticia Amadio -  Adecuacion a Autogestion	


Dim l_evaevenro
Dim l_evaevedesabr
Dim l_rs
Dim l_sql
' Variables
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Orden
  l_Orden     = "Empleado:;Apellido:;Nombre:"
  l_CamposOr  = "empleg;terape;ternom"

' Filtro
 l_Etiquetas  = "Empleado:;Apellido:;Nombre:"
 l_Campos     = "empleg;terape;ternom"
 l_Tipos      = "N;T;T"

l_evaevenro = request.querystring("evaevenro")  
l_evaevedesabr = request.querystring("evaevedesabr")  
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Formulario de Carga - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function seleccionar(){
ternro=document.ifrm.datos.cabnro.value;
empleg=document.ifrm.datos.empleg.value;
terape=document.ifrm.datos.terape.value;
ternom=document.ifrm.datos.ternom.value;
if ((ternro!=="") &&(ternro!=="0"))
	window.opener.nuevoempleado(ternro,empleg,terape,ternom);
window.close();
}

function agregarfiltro(){
	return " evaevenro=<%= l_evaevenro %> ";
}
function filtro(pag)
{
  abrirVentana('form_carga_eva_99.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
 // abrirVentana('/ess/ess/shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+ document.ifrm.datos.orden.value,'',250,160);
}

function orden(pag)
{
  abrirVentana(encodeURI('orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value),'',350,160)

	//abrirVentana('/ess/ess/shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+document.ifrm.datos.filtro.value,'',350,160)
  
}
</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="" method="post">
<table border="0" cellpadding="0" cellspacing="0">
  <tr style="border-color :CadetBlue;">
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
	        	<th align="left" class="th2"><%if ccodelco=-1 then%>Supervisados<%else%>Empleados<%end if%> en el Evento</th>
        		<td colspan="2" align="right" class="th2" valign="middle">
				<a class=sidebtnSHW href="Javascript:filtro('/ess/ess/eval/form_carga_eva_07.asp');">Filtro</a>
				<a class=sidebtnSHW href="Javascript:orden('form_carga_eva_07.asp');">Ordenar</a>
				&nbsp;
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td nowrap align="center" colspan="2">
		<br>
		<b>Evento de Evaluaci&oacute;n:</b>
		<input type="text" name="evaevedesabr" size="30" maxlength="30" value="<%= l_evaevedesabr %>" style="background : #e0e0de;">
	</td>
</tr>
<tr>
	<td align="right" colspan="2">
		<br>
		<b>Total:</b>
		<input readonly type="text" name="total" size="5" maxlength="5" value="0" style="background : #e0e0de;">
		&nbsp;&nbsp;&nbsp;
	</td>

</tr>
<tr>
	<td colspan="2" >
		<iframe name="ifrm" src="form_carga_eva_07.asp?filtro=evaevenro=<%= l_evaevenro %>" width="100%" height="290"></iframe> 
	</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:seleccionar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>	
</body>
</html>
