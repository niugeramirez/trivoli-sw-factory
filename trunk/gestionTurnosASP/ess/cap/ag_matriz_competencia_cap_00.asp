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
Archivo: ag_matriz_competencia_cap_00.asp
Descripción: 
Autor : Raul Chinestra
-->
<% 

on error goto 0

Dim l_sql
Dim l_rs
Dim rs9
Dim l_terape 
Dim l_ternom 
Dim l_terape2 
Dim l_ternom2 
Dim l_empleg
Dim l_puenro

Dim l_empfoto
Dim rs1
Dim sql
Dim l_orinro
Dim l_estnro

Dim l_ternro

Dim siguiente
Dim Anterior

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_orinro = request.querystring("origen")

if l_orinro = "" then
	l_orinro= "0"
end if

l_estnro = request.querystring("estado")

if l_estnro = "" then
	l_estnro= "-1" 'Pendientes
end if

'l_estado = request.querystring("estado")
l_orden = request.querystring("orden")
if l_orden = "" then
	l_orden= "empleg"
end if

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
  l_etiquetas = "C&oacute;digo:;Descripción:"
  l_Campos    = "evafactor.evafacnro;evafactor.evafacddesabr"
  l_Tipos     = "N;T"

' Orden
  l_Orden     = "C&oacute;digo:;Descripción:"
  l_CamposOr  = "evafactor.evafacnro;evafactor.evafacddesabr"

%>

<%

Dim l_convnro

' Se busca al empleado
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT ternro, terape, ternom, terape2, ternom2, empleg, empfoto "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE ternro=" & l_ternro
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0
l_terape = l_rs("terape")
l_ternom = l_rs("ternom")
l_terape2 = l_rs("terape2")
l_ternom2 = l_rs("ternom2")
l_empleg  = l_rs("empleg")
l_empfoto = trim(l_rs("empfoto") & " ")
l_ternro = l_rs("ternro")
l_rs.Close

Dim NombrePuesto

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

dim salir 

%>

<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Matriz de Competencias - Capacitación - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/menu_def.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script>

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString());
}

function Sig_Ant(leg)
{
if (leg != "")	{
	document.location ="ag_matriz_competencia_cap_00.asp?empleg=" + leg;
	}
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

function ActualizarGap(){
	document.ifrm.location ="gap_dina_modulos_cap_01.asp?ternro=" + document.datos.ternro.value;
}	

function orden(pag)
{
  alert(pag);
  abrirVentana('../shared/asp/orden_param_adp_00.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag)
{
  abrirVentana('../shared/asp/filtro_param_adp_00.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}


function param()
{
    alert('<%= l_ternro %>');
	var setear = "ternro=<%= l_ternro %>";
	return setear;
}


function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("gap_dina_modulos_cap_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value) + "&ternro=" + document.datos.ternro.value+ '&puenro= <%= l_puenro %>','execl',250,150);
}


function GenGap(){ 
  if (document.ifrm.datos.listanro.value == ''){
  	alert('Debe Seleccionar al menos un Módulo');
  }
  else {
  	abrirVentana("gap_dina_modulos_cap_04.asp?ternro=" + document.datos.ternro.value + "&modulos= " + document.ifrm.datos.listanro.value,'',550,150);
  }
}

function ActualizarGap(){
	document.ifrm1.location ="ag_matriz_competencia_cap_02.asp?ternro=" + document.datos.ternro.value + "&estado=" + document.datos.estnro.value;
}

function llamadaexcelmat(){ 
	if (filtro == "")
		Filtro(true);
	else
	//alert('Hacer ventana de excel 1');
		abrirVentana("ag_matriz_competencia_cap_excel_mat.asp?ternro=<%= l_ternro %>&orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'excel',250,150);
}
function llamadaexcelgap(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("ag_matriz_competencia_cap_excel_gap.asp?ternro=<%= l_ternro %>&estado=" + document.datos.estnro.value,'excel',250,150);
}

function Eventos(){
  if (document.ifrm1.datos.cabnro.value == '0'){
  	alert('Debe Seleccionar una Competencia.');
  }
  else {
	abrirVentana("ag_con_eventos_en_competencia_cap_00.asp?ternro=" + document.datos.ternro.value + "&competencia=" + document.ifrm1.datos.cabnro.value,'',750,400);
  }
}



// menu excel
HM_Array2 = [
[100,      // menu width
"mouse_x_position",
"mouse_y_position",
jsfont_color,   // font_color
jsmouseover_font_color,   // mouseover_font_color
'#FF7F50',   // background_color
'#FFA684',   // mouseover_background_color
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
["Matriz de Competencias","Javascript:llamadaexcelmat();",1,0,0],
["Gap Registrados de Comp.","Javascript:llamadaexcelgap();",1,0,0],
]



</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="" method="post">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<input type="hidden" name="empleg" value="<%= l_empleg %>">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
  <tr style="border-color :CadetBlue;">
     <th colspan="2" align="left">
		GAP
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
   		<td colspan="3" align="left" class="th2">
		<table width="100%">
			<tr>
				<td align="left" class="th2" width="100%"><b>Matriz de Competencias</b></td>
				<td class="th2" align="right" nowrap> 
					  &nbsp;&nbsp;&nbsp;						
				</td>						
			</tr>
		</table>		
	</td>								
   </tr>		
   <tr valign="top" height="45%">
      <td colspan="3" style="">
    	  <iframe  scrolling="Yes" name="ifrm" src="ag_matriz_competencia_cap_01.asp?ternro=<%= l_ternro %>" frameborder="0" src="blanc.asp" width="100%" height="100%"></iframe> 
	  </td>
   </tr>
   <tr>
     		<td colspan="3" align="left" class="th2">
				<table width="100%">
					<tr>
					    <td align="left" class="th2" width="100%"><b>GAP Registrado de Competencias</b></td>
						<td class="th2" align="right"><b>Estado:</b></td>
						<td class="th2">
					        <select style="width:160px" name=estnro size="1" onchange="Javascript:ActualizarGap()">
						       <option value="2">Todos</option>
						       <option value="-1">Pendiente</option>
						       <option value="0">Terminado</option>
					        </select>
					        <script>document.datos.estnro.value= <%= l_estnro %></script>		
				         </td>	
			            <td class="th2">&nbsp;</td>
						<td class="th2" align="right" nowrap> 
							  <a class=sidebtnABM href="Javascript:Eventos();">Eventos</a>
		                      <% 'call MostrarBoton ("sidebtnSHW", "Javascript:abrirVentana('gap_dina_competencias_cap_00.asp?empleg="& l_empleg &"' + '&puenro="& l_puenro &"','',720,480);","Gap Dinámicos")%>
							  &nbsp;&nbsp;&nbsp;						
						</td>						
					</tr>
				</table>		
			</td>								
		</tr>							
   <tr valign="top" height="50%">
     <td colspan="3" style="">
    	  <iframe scrolling="Yes" frameborder="0" name="ifrm1" src="ag_matriz_competencia_cap_02.asp?ternro=<%= l_ternro %>&estado=<%= l_estnro %>" width="100%" height="100%"></iframe> 
	 </td>
   </tr>		
</table>
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>

