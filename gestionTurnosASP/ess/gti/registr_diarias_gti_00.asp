<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
---------------------------------------------------------------------------------------------
Archivo: registr_diarias_gti_00.asp
Descripcion: Modulo que se encarga de ABM de registraciones.
Modificacion:
   23/07/2003 - Scarpa D. - Permitir seleccionar un empleado
   11/08/2003 - Scarpa D. - Al seleccionar ver todos los empleados (Activos/Inactivos)
   14/08/2003 - Scarpa D. - Agregado del boton de depuracion
   20/08/2003 - Scarpa D. - Llamar al 09 en vez del 08 cuando se realiza un cambio de estado.   
   18/09/2003 - Scarpa D. - No mostrar los botones anterior y siguiente cuando se llama desde 
                            el tablero de rango de fechas.
   18/02/2004 - Alvaro Bayon - Para borrar una registración, esta no debe ser manual y además el usuario
   								debe tener perfil de supervisor
   18/02/2004 - Scarpa D.    - Para modificar una registración, esta no debe ser manual y además el usuario
   								debe tener perfil de supervisor
   29/09/2006 - Mariano Capriz - Se corrigio la validacion del mes para desde y hasta ya que estaba preguntando si el mes era menor a 9 enves de ser menor a 10
---------------------------------------------------------------------------------------------
-->

<% 
on error goto 0

 Dim l_tipo 
     l_tipo = request.queryString("tipo")
 Dim l_rs
 Dim l_rs1
 Dim l_sql
 
 Dim l_ternro
 Dim l_empleg
 Dim l_empleado
 Dim l_datos
 Dim l_fechadesde
 Dim l_fechahasta
 
' Para mss
 Dim l_emplegsec
 Dim l_emplogueado
 Dim l_ternrolog
 
' Variables
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
 Dim l_Orden      ' Son las etiquetas que aparecen en el orden
 Dim l_CamposOr   ' Son los campos para el orden
 
 
' Orden
 l_Orden     = "Descripci&oacute;n:;Tipo de Hora:;Fecha Desde:;Fecha Hasta:"
 l_CamposOr  = "gnovdesabr;gtnovdesabr;gnovdesde;gnovhasta"
 
 
 l_ternro 	  = l_ess_ternro
 l_empleg 	  = l_ess_empleg
 l_empleado   = request.querystring("empleado")
 l_fechadesde = request.querystring("fechadesde")
 l_fechahasta = request.querystring("fechahasta")
 
if l_fechadesde = "" then
   if month(Date()) < 10 then
		l_fechadesde = "01/0" & month(Date()) & "/" & year(date())
	else
		l_fechadesde = "01/" & month(Date()) & "/" & year(date())	
	end if
end if	

if l_fechahasta = "" then
   if month(Date()) = 12 then
      l_fechahasta = "31/12/" & year(date())
   else
      if month(Date()) < 9 then
         l_fechahasta = "01/" & (month(Date()) + 1) & "/" & year(date())
	  else
	     l_fechahasta = "01/0" & (month(Date()) + 1) & "/" & year(date())
	  end if
	  l_fechahasta = dateadd("d", -1, CDate(l_fechahasta) )
   end if
end if

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
 
 l_ternrolog = 0
 
 l_datos = "ternro=" & l_ternro & "&empleg=" & l_empleg & "&empleado=" & l_empleado
 l_datos = l_datos & "&fechadesde=" & l_fechadesde & "&fechahasta=" & l_fechahasta

%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Registraciones - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_help_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script src="fn_ay_empleado.js"></script>
<script>

HM_Array1 = [
[120,      // menu width
"mouse_x_position",
"mouse_y_position",
jsfont_color,   // font_color
jsmouseover_font_color,   // mouseover_font_color
'#CD5C5C',   // background_color
'#FFA07A',   // mouseover_background_color
'#ffffff',   // border_color
jsseparator_color,    // separator_color
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
["Reloj y Fecha","Javascript:orden(1)",1,0,0],
["Fecha y Hora","Javascript:orden(2)",1,0,0],
]

</script>

<script>

function orden(ord)
{
	refrescar(ord);
}

function filtro(pag)
{
  abrirVentana('filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.tipo_dias.datos.orden.value,'',250,160);
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

function completar2(str){
	dato = str + "";
	if (dato.length == 1){
		dato= "0" + dato;
	}	
	return dato;	
}
function setearfechas(){
	var fechanueva=new Date(); 
   	str =  completar2(fechanueva.getDate()) + "/";
  	str += completar2((fechanueva.getMonth() + 1)) + "/";
 	str += fechanueva.getYear();
   document.datos.fechadesde.value = str;
   document.datos.fechahasta.value = str;
}

function menorigual(fecha1,fecha2){
	var f1= new Date(); 
	f1.setFullYear(fecha1.substr(6,4),fecha1.substr(3,2)-1,fecha1.substr(0,2));
	var segf1=Date.parse(f1); 

	var f2= new Date(); 
	f2.setFullYear(fecha2.substr(6,4),fecha2.substr(3,2)-1,fecha2.substr(0,2));
	var segf2=Date.parse(f2); 

	if (segf1<=segf2){return true}
	else{return false}
}

function validar(){
	if (validarfecha(document.datos.fechadesde) && validarfecha(document.datos.fechahasta))
	{
		if (menorigual(document.datos.fechadesde.value,document.datos.fechahasta.value))
		return true
		else {
		alert("La fecha desde debe ser menor o igual que la fecha hasta.")	
		return false
		}
	}
}

function refrescar(ord){
var datos="";
if (validar()) {
	datos = 'empleg='+document.datos.empleg.value + '&fechadesde='+ document.datos.fechadesde.value + '&fechahasta='+ document.datos.fechahasta.value;
	document.ifrm.location="registr_diarias_gti_01.asp?" + datos + "&orden="+ ord
}
}

function filtro(){
var fil = "";
fil= "ternro="+document.datos.ternro.value + "&empleg=" + document.datos.empleg.value + "&empleado=" + document.datos.empleado.value;
fil= fil + "&fechadesde=" + document.datos.fechadesde.value + "&fechahasta=" + document.datos.fechahasta.value;

return fil;
}

function borrar(){
	//alert('La Operación SI está Permitida para el Perfil del Usuario.');
	eliminarRegistro(document.ifrm,'registr_diarias_gti_04.asp?cabnro=' + document.ifrm.datos.cabnro.value + '&' + filtro())
}

function noborrar(){
	alert('La Operación no está Permitida para el Perfil del Usuario.');
}

function validarBorrado(){
	if (document.ifrm.datos.cabnro.value == 0)
		{
		alert("Debe selecionar un registro para realizar la operación.");
		}
	else
		document.valida.location = "registr_diarias_gti_14.asp?regnro=" + document.ifrm.datos.cabnro.value
}

function validarModif(){
	if (document.ifrm.datos.cabnro.value == 0)
		{
		alert("Debe selecionar un registro para realizar la operación.");
		}
	else
		document.valida.location = "registr_diarias_gti_14.asp?regnro=" + document.ifrm.datos.cabnro.value + '&tipo=M';
}


function modificar(){
    abrirVentanaVerif('registr_diarias_gti_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value+ '&' + filtro(),'',470,190)
}


</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form name="datos" action="" method="post">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">
<input type="Hidden" name="empleg" value="<%= l_empleg %>">
<input type="Hidden" name="empleado" value="<%= l_empleado %>">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
  <tr style="border-color :CadetBlue;">
	<td colspan="2">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
	        	<th align="left">Registraciones</th>
        		<th colspan="2" align="right" valign="middle">
				<% call MostrarBoton("sidebtnABM", "Javascript:abrirVentana('registr_diarias_gti_02.asp?Tipo=A&'+ filtro(),'',470,190)", "Alta")%>
				<% call MostrarBoton("sidebtnABM", "Javascript:borrar();", "Baja")%>
				<% call MostrarBoton("sidebtnABM", "Javascript:abrirVentanaVerif('registr_diarias_gti_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value+ '&' + filtro(),'',470,190);", "Modifica")%>
				&nbsp;&nbsp;&nbsp;
				<a class=sidebtnSHW href="#" onClick="MenuPopUp('elMenu1',event)" onMouseOut="MenuPopDown('elMenu1')">Ordenar</a>
				</th>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" height="10">
		<br>
	</td>
</tr>
<tr>
	<td>
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td align="right">
				<b>Desde:</b>
				<input type="text" name="fechadesde" size="10" maxlength="10" value="<%= l_fechadesde %>">
				<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
				</td>
				<td align="right">
				<b>Hasta:</b>
				<input type="text" name="fechahasta" size="10" maxlength="10" value="<%= l_fechahasta %>">
				<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
				</td>
			</tr>
		</table>
	</td>
	<td> 
		&nbsp;&nbsp;
		<a class="sidebtnSHW" href="Javascript:;" onclick="Javascript:refrescar(2);">Actualizar</a>
		&nbsp;&nbsp;
	</td>
</tr>
<tr>
	<td colspan="2" height="10">
		<br>
	</td>
</tr>
<tr height="95%">
	<td colspan="2" >
		<iframe name="ifrm" src="registr_diarias_gti_01.asp?<%= l_datos %>" width="100%" height="100%"></iframe> 
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="0" height="0"></iframe> 
</form>	
<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>
</body>
</html>
