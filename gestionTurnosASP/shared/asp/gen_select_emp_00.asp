<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/gengrup.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_00.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Modulo que se encarga de la seleccion de empleados
Modificacion   :
  04/12/2003 - Scarpa D. - Indicar una seleccion inicial para los empleados no selectados
  15/12/2003 - Scarpa D. - Correccion al modificar la lista de empleados
  23/12/2003 - Scarpa D. - Poder setear el titulo de la ventana
  07/01/2004 - Scarpa D. - Correccion por problemas en los filtros derecho e izquierdo
  09/01/2004 - Scarpa D. - Restringir el conjunto de empleados que puede aparecer en el lado izq.
-----------------------------------------------------------------------------
-->
<%
Dim l_srcDatos   
Dim l_funtion    
Dim l_ventana    
Dim l_formulario 
Dim l_funcCerrar
Dim l_seltipnro
Dim l_filtroIni
Dim l_filtroOnly
Dim l_titulo

Dim l_selalto
Dim l_selancho

l_srcDatos   = request("srcdatos")
l_funtion    = request("funcion")
l_ventana    = request("ventana")
l_formulario = request("formulario")
l_funcCerrar = request("funccerrar")
l_seltipnro  = request("seltipnro")
l_filtroIni  = request("filtroIni")
l_titulo     = request("titulo")
l_filtroOnly = request("filtroOnly")

l_selalto    = request("selalto")
l_selancho   = request("selancho")

if l_selalto = "" then
   l_selalto = "280"
end if

if l_selancho = "" then
   l_selancho = "300"
end if

if l_titulo = "" then
   l_titulo = "Selecci&oacute;n"
end if

%>
<script src="/turnos/shared/js/fn_windows.js"></script>
<SCRIPT SRC="/turnos/shared/js/menu_def.js"></SCRIPT>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>

HM_Array1 = [
[120,      // menu width
"mouse_x_position",
"mouse_y_position",
jsfont_color,   // font_color
jsmouseover_font_color,   // mouseover_font_color
jsbackground_color,   // background_color
jsmouseover_background_color,   // mouseover_background_color
jsborder_color,   // border_color
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
["Por empleado","Javascript:ordenar(1)",1,0,0],
["Por apellido","Javascript:ordenar(0)",1,0,0],
]

var lado = 1; //1 = izq, otro = der
var MaxSQL = 1000;
var arrSQL = new Array(MaxSQL);
var i;

for(i=0; i <MaxSQL;i++){
  arrSQL[i] = '';
}

function filtroIz()
{
	 var jsNuevo = Nuevo_Dialogo(window, "../../adp/dialog_frame.asp?titulo=Filtro seleccion&pagina=gengrup_v3_filtro.asp", 30, 35);
	 	
	 if ((jsNuevo != null) && (jsNuevo != "")){
		document.datos.sqlfiltroizq.value = jsNuevo.substr(0);
        actualizar();		
	 }  
}

function filtroDe()
{
	 var jsNuevo = Nuevo_Dialogo(window, "../../adp/dialog_frame.asp?titulo=Filtro seleccion&pagina=gengrup_v3_filtro.asp", 30, 35);
	 	
	 if ((jsNuevo != null) && (jsNuevo != "")){
		document.datos.sqlfiltroder.value = jsNuevo.substr(0);
        actualizarSelectados();
	 }  
}

function ordenar(orden){
	if (lado == 1){
	   if (orden==1)
		  document.datos.sqlordenizq.value = "ORDER BY empleg";
	   else  
		  document.datos.sqlordenizq.value = "ORDER BY terape, ternom";	
		  
       actualizar();		  
	}else{
	   if (orden==1)
		  document.datos.sqlordender.value = "ORDER BY empleg";
	   else  
		  document.datos.sqlordender.value = "ORDER BY terape, ternom";	

       actualizarSelectados();
	}
}

/* Se fija si existe una clave en la lista */
function existe(cabnro){
  var arreglo;
  var nro;
  var pos=0;
  var lista= document.datos.seleccion.value;
  
  if ((lista!="") && (lista!=null)){
      arreglo = lista.split(',');
	  while (pos < arreglo.length){
	     nro = arreglo[pos];
		 pos++;
		 if (nro == cabnro){
		   return 1;
		 }
	  }
  }
  return 0;  
}

/* Agrega una clave a la lista si no existe */
function agregalista(cabnro){
  if (!existe(cabnro)){
    if (document.datos.seleccion.value == ''){
       document.datos.seleccion.value = cabnro	
	}else{
  	   document.datos.seleccion.value = document.datos.seleccion.value + ',' + cabnro
	}    
  }
}

function quitalista(cabnro)
{
  var arreglo;
  var nro;
  var pos=0;
  var listatmp = "";
  var lista= document.datos.seleccion.value;
  
  if ((lista!="") && (lista!=null)){
	  arreglo = lista.split(',');
	  while (pos < arreglo.length){
	     nro = arreglo[pos];
		 pos++;
		 if (nro != cabnro){
			 if (listatmp == ''){
		  	   listatmp = nro;
			 }else{
			   listatmp = listatmp + ',' + nro;  
			 }
		 }
	  }
      document.datos.seleccion.value = listatmp;	  
  }
  
}

function Todos(fuente,destino,totorig,totdest)
{
	x=fuente.length;
	for (i=1;i<=x;i++)
	{
		var opcion = new Option();
		opcion.value = fuente[0].value;
		opcion.text  = fuente[0].text;
		fuente.remove(0);
		destino.add(opcion);
		if (fuente.name == 'nselfil')
		   agregalista(opcion.value);
		else   
		   quitalista(opcion.value);
		totorig.value = Number(totorig.value) - 1;
		totdest.value = Number(totdest.value) + 1;
	}
    document.datos.filtroder.value = selfil.registro.selfil.length;
    document.datos.filtroizq.value = nselfil.registro.nselfil.length;
}

function Uno(fuente,destino,totorig,totdest)
{
    if (fuente.selectedIndex != -1) {
	    var opcion = new Option();
	    opcion.value= fuente[fuente.selectedIndex].value;
	    opcion.text  = fuente[fuente.selectedIndex].text;
	    fuente.remove(fuente.selectedIndex);
	    destino.add(opcion);
		if (fuente.name == 'nselfil')
		   agregalista(opcion.value);
		else   
		   quitalista(opcion.value);
	    destino[destino.length-1].focus();
	    document.datos.filtroder.value = selfil.registro.selfil.length;
	    document.datos.filtroizq.value = nselfil.registro.nselfil.length;
		totorig.value = Number(totorig.value) - 1;
		totdest.value = Number(totdest.value) + 1;
	}
}

function Lista()
{
     return document.datos.seleccion.value;
}

function Aceptar(srcDatos,funcion,ventana,formulario)
{

    if (srcDatos != ''){
	   var obj = eval(srcDatos);
	    obj.value = Lista();
	}
	
	if (funcion != ''){
	   eval(funcion + '()');
	}
    
	if (ventana != ''){
	   abrirVentana(ventana,'',200,200);
	}
	
	if (formulario != ''){
	   abrirVentanaH('','voculta',200,200);
       document.datos.target = 'voculta';
       document.datos.action = formulario;
       document.datos.submit();	   
	}	
	
    window.close();
}

function cambioSQL(selnro,sql){
   arrSQL[parseInt(selnro)] = sql;  
   actualizar();
}

function actualizar(){   
   var arrlista;

   if (document.ifrmfiltros.datos.listanro.value != "")   
       arrlista = document.ifrmfiltros.datos.listanro.value.split(',');
   else
       arrlista = "";	   
   
   document.datos.sqlfiltroemp.value = '';

   for (i=0; i < arrlista.length; i++){
       if (document.datos.sqlfiltroemp.value == ''){
          document.datos.sqlfiltroemp.value = arrSQL[parseInt(arrlista[i])];	   
	   }else{
          document.datos.sqlfiltroemp.value = document.datos.sqlfiltroemp.value + ';' + arrSQL[parseInt(arrlista[i])];
	   }
   }

   document.datos.target = 'nselfil';
   document.datos.action = 'gen_select_emp_02.asp';
   document.datos.submit();
}

function  actualizarSelectados(){
   document.datos.target = 'selfil';
   document.datos.action = 'gen_select_emp_03.asp';
   document.datos.submit();
}  

function altaCriterio(){
   abrirVentana('criterios_02.asp?tipo=A','',470,380);
}

function bajaCriterio(){
   var arreglo = document.ifrmfiltros.datos.listanro.value.split(',');
	
   if ((arreglo != 0) && (arreglo.length == 1)){
      if (parseInt(document.ifrmfiltros.deSistema(arreglo[0])) == 0){
         if (confirm('¿Desea borrar el criterio?.')){
           abrirVentanaH('criterios_04.asp?selnro='+ arreglo[0] ,'',470,380);		 
	     }
	  }else{
	     alert('No se puede eliminar un registro del sistema.');
	  }	 
   }else{
      alert('Debe seleccionar un criterio.');
   }	
}

function modifCriterio(){
   var arreglo = document.ifrmfiltros.datos.listanro.value.split(',');
	
   if ((arreglo != 0) && (arreglo.length == 1)){
      if (parseInt(document.ifrmfiltros.deSistema(arreglo[0])) == 0){   
         abrirVentana('criterios_02.asp?tipo=M&selnro='+ arreglo[0] ,'',470,380);		 
	  }else{
	     alert('No se puede modificar un registro del sistema.');
	  }	 	  
   }else{
      alert('Debe seleccionar un criterio.');
   }	
}

function agregarCriterio(){
   abrirVentanaCent('','ventcriterio',470,380);
   document.datos.target = 'ventcriterio';
   document.datos.action = 'criterios_02.asp?tipo=LE';
   document.datos.submit();   
}

function cerrar(){
  if ('<%= l_funcCerrar%>' != ''){
     eval('<%= l_funcCerrar%>' + '()');
  }
  window.close();
}

</script>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%><%= l_titulo%></title>
</head>
<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">

<form name="datos" method="post" action="" target="">
<input type="Hidden" value="0" name="seleccion">
<input type="Hidden" name="sqlfiltroemp"  value="<%= l_filtroIni%>">
<input type="Hidden" name="sqlfiltroonly" value="<%= l_filtroOnly%>">
<input type="Hidden" value="" name="sqlfiltroizq">
<input type="Hidden" value="" name="sqlfiltroder">
<input type="Hidden" value="" name="sqlordenizq">
<input type="Hidden" value="" name="sqlordender">
<input type="Hidden" value="<%= l_selalto%>"  name="selalto">
<input type="Hidden" value="<%= l_selancho%>" name="selancho">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td class="th2" colspan="2" height="2%">
	<script>document.write(document.title);</script>
	</td>	
    <td align="right" class="th2">
	  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
<td colspan="3" align="center" height="2%">
  <table style="width:550px;border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">
    <tr>
	  <td align="center">
         <b>Criterios de Selecci&oacute;n</b><br>
         <iframe frameborder="1" name="ifrmfiltros" src="gen_select_emp_01.asp?seltipnro=<%= l_seltipnro%>" width="98%" height="100" scroll=no></iframe><br>
	     <input type="Radio" checked name="sqloperando" value="AND" onclick="javascript:actualizar();">Considerar todos los criterios.&nbsp;&nbsp;&nbsp; 
	     <input type="Radio" name="sqloperando" value="OR" onclick="javascript:actualizar();">Considerar al menos un criterio.  	
	  </td>	
	  <td align="left" valign="middle">
	     <a class=sidebtnSHW style="width:75px" href="javascript:altaCriterio();">Alta</a><br>
	     <a class=sidebtnSHW style="width:75px" href="javascript:bajaCriterio();">Baja</a><br>
	     <a class=sidebtnSHW style="width:75px" href="javascript:modifCriterio();">Modificaci&oacute;n</a>		 
	  </td>
	</tr>
  </table>		
</td>
</tr>
<tr>
  <td colspan="3" align="center">
    <table style="border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">    
	  <tr>
  
    <td align="left" width="45%" valign="top" height="2%">
	<a class=sidebtnSHW href="javascript:filtroIz();">Filtro</a>
	<a class=sidebtnSHW href="#" onClick="lado=1; MenuPopUp('elMenu1',event)" onMouseOut="lado=1; MenuPopDown('elMenu1')">Orden</a><br><br>
	<b>No Seleccionados</b><br>
	Filtro:&nbsp;
    <input type="Text" name="filtroizq" size="6" class="hidden" readonly>
	Total:&nbsp;   <input type="Text" name="totalizq" size="6" class="hidden" readonly>
	</td>
    <td align=left width="10%">
    </td>
    <td width="45%" align="left" valign="top" height="2%">
	<a class=sidebtnSHW href="javascript:filtroDe();">Filtro</a></a>
	<a class=sidebtnSHW href="#" onClick="lado=0; MenuPopUp('elMenu1',event)" onMouseOut="lado=0; MenuPopDown('elMenu1')">Orden</a>
	<a class=sidebtnSHW href="javascript:agregarCriterio();">Agregar Criterio</a></a><br><br>	
	<b>Seleccionados</b><br>
	Filtro:&nbsp;
    <input type="Text" name="filtroder" size="6" class="hidden" readonly>
    Total:&nbsp;
    <input type="Text" name="totalder" size="6" class="hidden" readonly>
	</td>
</tr>
<tr>
  <td valign="top" align="left">
    <iframe align="center" name="nselfil" width="<%= l_selancho%>" height="<%= l_selalto%>" src="" scroll=no></iframe>   
  </td>
  <td valign="middle" align="center" width="10%" >
	<a class=sidebtnSHW href="javascript:Todos(nselfil.registro.nselfil, selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);">>></a>
	<a class=sidebtnSHW href="javascript:Uno(nselfil.registro.nselfil,selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);">></a>
	<a class=sidebtnSHW href="javascript:Uno(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);"><</a>
	<a class=sidebtnSHW href="javascript:Todos(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);"><<</a>
  </td>  
  <td valign="top" align="left">
    <iframe align="center" name="selfil" width="<%= l_selancho%>" height="<%= l_selalto%>" src=""></iframe>   
  </td>
</tr>    

</table>
</td>
</tr>
<tr>
    <td align="right" class="th2" colspan="3" height="2%">
		<a class=sidebtnABM href="javascript:Aceptar('<%= l_srcDatos%>','<%= l_funtion%>','<%= l_ventana%>','<%= l_formulario%>')">Aceptar</a>
		<a class=sidebtnABM href="Javascript:cerrar();">Cancelar</a>
	</td>
</tr>
</table>
</form>

<script>
  document.datos.seleccion.value = eval('<%= l_srcDatos & ".value"%>');
  if (document.datos.seleccion.value == ""){
      document.datos.seleccion.value = '0';
  }
  actualizarSelectados();
  document.datos.target = 'nselfil';
  document.datos.action = 'gen_select_emp_02.asp';
  document.datos.submit();  
</script>

<SCRIPT SRC="/turnos/shared/js/menu_op.js"></SCRIPT>

</html>

