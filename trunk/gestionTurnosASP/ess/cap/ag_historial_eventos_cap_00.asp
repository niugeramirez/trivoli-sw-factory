<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/antigfec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo		: ag_historial_eventos_cap_00.asp
Descripcion	: Consulta del historial de Eventos 
Autor		: Gustavo Ring	
Fecha		: 07/06/2007
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden

' Filtro
  l_etiquetas = "C&oacute;d. Ext.:;Descripci&oacute;n:;Curso:"
  l_Campos    = "evecodext;evedesabr;curdesabr"
  l_Tipos     = "T;T;T"

' Orden
  l_Orden     = "C&oacute;d. Ext.:;Descripci&oacute;n:;Curso:"
  l_CamposOr  = "evecodext;evedesabr;curdesabr"

Dim l_orden2
Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_terape2 
Dim l_ternom2 
Dim l_empleg
Dim l_modulo
Dim l_puenro
Dim l_convnro

Dim l_empfoto
Dim rs1
Dim sql

Dim l_ternro

Dim siguiente
Dim Anterior

l_ternro = l_ess_ternro

l_orden2 = request.querystring("orden")
if l_orden2 = "" then
	l_orden2 = "empleg"
end if

' Se busca al empleado
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT terape, ternom, terape2, ternom2, empleg, ternro "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE empleg=" & l_ess_empleg

rsOpen l_rs, cn, l_sql, 0

l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
l_terape2 = l_rs("terape2")
l_ternom2 = l_rs("ternom2")
l_ternro  = l_rs("ternro")

l_rs.Close
%>

<html>
<head>
<link href="../<%= c_Estilo %>" rel="StyleSheet" type="text/css">
<title>Participación en cursos abiertos- Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/menu_def.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>

function param()
{
	var setear = "ternro=<%= l_ternro %>";
	return setear;
}

function orden(pag)
{
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}


function Tecla(num){
  if (num==13) {
		verificacodigo(document.datos.empleg,document.datos.empleado,'empleg','terape, ternom','empleado');
		Sig_Ant(document.datos.empleg.value);
		return false;
  }
  return num;
}

function emplerror(nro){
	alert('empleado error:'+nro);
}

function datos(){
var dat='';
 dat='ternro='+document.datos.ternro.value+"&empleg="+document.datos.empleg.value+'&empleado='+document.datos.empleado.value;
 dat= dat+ '&fechadesde='+document.datos.fecha.value+ '&fechahasta='+document.datos.fecha.value;
 return dat;
}
function nuevoempleado(ternro,empleg,terape,ternom)
{
if (empleg != 0) 
	{
	document.datos.empleg.value = empleg;
	document.datos.empleado.value = terape + ", " + ternom;
	Sig_Ant(document.datos.empleg.value);
	}
}

function llamadaexcel(){ 
	abrirVentana("ag_eventos_abiertos_cap_excel.asp?ternro=<%= l_ternro%>&orden="+document.ifrm.datos.orden.value+"&filtro="+escape(document.ifrm.datos.filtro.value),'execl',50,50);
}

function candidato(){
	if (document.ifrm.datos.cabnro.value == 0){
		alert("Debe seleccionar un Evento.");
	}else{
		//alert("ok");
		abrirVentanaH('ag_eventos_abiertos_cap_03.asp?ternro=<%= l_ternro %>&evenro=' + document.ifrm.datos.cabnro.value + '&tipo=C','','','');
		//abrirVentanaH('participantes_cap_02.asp?ternro=<%'= l_ternro %>&evenro=' + document.ifrm.datos.cabnro.value + '&tipo=C','','','');
	}
}

function participante(){
	if (document.ifrm.datos.cabnro.value == 0){
		alert("Debe seleccionar un Evento.");
	}else{
		 if ((ifrm.jsSelRow.cells(4).innerText.slice(0) < 1)) {
			alert("No queda lugar para Participantes en el evento ");
		 }else{
			//alert(document.ifrm.datos.cabnro.value);
			abrirVentanaH('ag_eventos_abiertos_cap_03.asp?ternro=<%= l_ternro %>&evenro=' + document.ifrm.datos.cabnro.value + '&tipo=P','','','');
			//abrirVentanaH('participantes_cap_02.asp?ternro=<%'= l_ternro %>&evenro=' + document.ifrm.datos.cabnro.value + '&tipo=P','','','');
		}
	}
}

function quitar(){
	if (document.ifrm2.datos.cabnro.value == 0){
		alert("Debe seleccionar un Evento.");
	}else{
		//alert("ok");
		abrirVentanaH('ag_eventos_abiertos_cap_03.asp?ternro=<%= l_ternro %>&evenro=' + document.ifrm2.datos.cabnro.value + '&tipo=Q','','','');
	}
}

function quitar2(){
	if (document.ifrm.datos.cabnro.value == 0){
		alert("Debe seleccionar un Evento.");
	}else{
		//alert("ok");
		abrirVentanaH('ag_eventos_cerrados_cap_03.asp?ternro=<%= l_ternro %>&evenro=' + document.ifrm.datos.cabnro.value + '&tipo=Q','','400','150');
	}
}

//MENU: eventos
HM_Array1 = [
[110,      // menu width
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
["Candidato","Javascript:candidato();",1,0,0],
["Participante","Javascript:participante();",1,0,0],
]


</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" >

<form name="datos" action="" method="post">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<%

dim salir 

%>

<table border="0" cellpadding="0" cellspacing="0"  height="100%" width="100%">

<tr>
	<td colspan="2" height="100%">
		<table border="0" cellpadding="0" cellspacing="0"  height="100%" width="100%">
			<tr>
				<th>Actividades de Capacitación Realizadas.</th>
			</tr>
	        <tr valign="top">
	        	<td align="left" style="" height="100%" nowrap  colspan="2">
   	  				<iframe  scrolling="Yes" name="ifrm2" src="ag_historial_eventos_cap_01.asp?ternro=<%= l_ternro %>"></iframe> 
   				</td>        				
   			</tr>
		</table>
	</td>
</tr>
</form>	
</table>

<% 
cn.Close
set cn = Nothing
%>

<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
<script>
//window.document.body.scroll = "no";
</script>
</body>
</html>
