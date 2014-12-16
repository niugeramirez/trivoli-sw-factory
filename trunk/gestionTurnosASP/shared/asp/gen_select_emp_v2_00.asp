<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/gengrup.inc"-->
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
Dim l_titulo
Dim l_filtroOnly

Dim l_selalto
Dim l_selancho
Dim l_vent_pagina

l_vent_pagina = 100

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
   l_selalto = "265"
end if

if l_selancho = "" then
   l_selancho = "300"
end if

if l_titulo = "" then
   l_titulo = "Selecci&oacute;n"
end if

%>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<SCRIPT SRC="/serviciolocal/shared/js/menu_def.js"></SCRIPT>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
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

function abrirDialogo(ventana,archivo,ancho,alto){
   return (showModalDialog(archivo,'','dialogWidth:' + ancho +';dialogHeight:' + alto + ';help: 0; status: 0; resizable:0; center:1;scroll:0'));
}
	
function filtroIz()
{
	 var jsNuevo = abrirDialogo(window, "../../adp/dialog_frame.asp?titulo=Filtro seleccion&pagina=gengrup_v5_filtro.asp", 30, 20);
	 	
	 if ((jsNuevo != null) && (jsNuevo != "")){
		document.datos.sqlfiltroizq.value = jsNuevo.substr(0);
		document.datos.paginaizq.value = 1;
        actualizar();		
	 }  
}

function filtroDe()
{
	 var jsNuevo = abrirDialogo(window, "../../adp/dialog_frame.asp?titulo=Filtro seleccion&pagina=gengrup_v5_filtro.asp", 30, 20);
	 	
	 if ((jsNuevo != null) && (jsNuevo != "")){
		document.datos.sqlfiltroder.value = jsNuevo.substr(0);
		document.datos.paginader.value = 1;		
        actualizarSelectados();
	 }  
}

function ordenar(orden){
	if (lado == 1){
	   if (orden==1)
		  document.datos.sqlordenizq.value = "ORDER BY empleg";
	   else  
		  document.datos.sqlordenizq.value = "ORDER BY terape, ternom";	

 	   document.datos.paginaizq.value = 1;		  
       actualizar();		  
	}else{
	   if (orden==1)
		  document.datos.sqlordender.value = "ORDER BY empleg";
	   else  
		  document.datos.sqlordender.value = "ORDER BY terape, ternom";	

	   document.datos.paginader.value = 1;
       actualizarSelectados();
	}
}

function Todos(fuente,destino,totorig,totdest)
{
    var listaN='';
   
	x=fuente.length;

	for (i=1;i<=x;i++)
	{
	    if (listaN == ''){
		    listaN = fuente[0].value;
		}else{
		    listaN = listaN + ',' + fuente[0].value;			
		}
	
		if (fuente.name != 'nselfil'){
			var opcion = new Option();
			opcion.value = fuente[0].value;
			opcion.text  = fuente[0].text;
			destino.add(opcion);
		}
		fuente.remove(0);
	}

	if (fuente.name == 'nselfil')
	   document.datos.accion.value = "A";
	else   
       document.datos.accion.value = "Q";

	if (Trim(listaN) != ''){	
	    document.datos.listanueva.value = listaN;
	
	    document.datos.target = 'valida';
	    document.datos.action = 'gen_select_emp_v2_04.asp';
	    document.datos.submit();
		
		if (document.ifrmfiltros.datos.listanro.value == ""){
		    document.datos.filtroizq.value = nselfil.registro.nselfil.length;
			if (fuente.name == 'nselfil')
			   totorig.value = nselfil.registro.nselfil.length;
			else   
			   totdest.value = nselfil.registro.nselfil.length;
		}
	}
}

function Uno(fuente,destino,totorig,totdest){

    var listaN='';
	var selFuente = fuente.selectedIndex;
	var selDestino = ",";
	while (fuente.selectedIndex != -1 ){
		selDestino = selDestino + fuente[fuente.selectedIndex].value + ",";	
		
	    if (fuente.selectedIndex != -1) {
		    if (listaN == ''){
			    listaN = fuente[fuente.selectedIndex].value;
			}else{
			    listaN = listaN + ',' + fuente[fuente.selectedIndex].value;			
			}
			
			if (fuente.name != 'nselfil'){
	            var opcion = new Option();
	      	    opcion.value= fuente[fuente.selectedIndex].value;
			    opcion.text  = fuente[fuente.selectedIndex].text;
			    destino.add(opcion);
			}
            fuente.remove(fuente.selectedIndex);			
		}
	}
	
	if (fuente.name == 'nselfil')
	   document.datos.accion.value = "A";
	else   
       document.datos.accion.value = "Q";
	   
	if (Trim(listaN) != ''){
	    document.datos.listanueva.value = listaN;
	
	    document.datos.target = 'valida';
	    document.datos.action = 'gen_select_emp_v2_04.asp';
	    document.datos.submit();

		if (document.ifrmfiltros.datos.listanro.value == ""){
		    document.datos.filtroizq.value = nselfil.registro.nselfil.length;
			if (fuente.name == 'nselfil')
			   totorig.value = nselfil.registro.nselfil.length;
			else   
			   totdest.value = nselfil.registro.nselfil.length;		
		}
    }		
	
}

// Setea el focus en un items (codigo) del select (objeto).
function Reposicionar (objeto, codigo){
	objeto.selectedIndex = -1;
	for (i=0;i<objeto.length;i++){
		if (codigo.indexOf(","+objeto[i].value+",") != -1){
			objeto[i].selected = true;
			objeto[i].focus;
		}
	}
}


function Lista()
{
     return document.datos.seleccion.value;
}

function Aceptar(srcDatos,funcion,ventana,formulario)
{
//    obtenerTerceros();
    document.all.botAceptar.disabled = true;
	document.all.botAceptar.className = "sidebtnDSB"
	
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
   document.datos.paginaizq.value = 1;   
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
   document.datos.action = 'gen_select_emp_v2_02.asp';
   document.datos.submit();
}

function  actualizarSelectados(){
   document.datos.target = 'selfil';
   document.datos.action = 'gen_select_emp_v2_03.asp';
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

function sigPagNoSel(){
   if (document.datos.paginaIzqFin.value == "0"){
      document.datos.paginaizq.value = parseInt(document.datos.paginaizq.value,10) + 1;
      actualizar();
   }
}

function antPagNoSel(){
   if (parseInt(document.datos.paginaizq.value,10) > 1){
      document.datos.paginaizq.value = parseInt(document.datos.paginaizq.value,10) - 1;
      actualizar();
   }
}

function sigPagSel(){
   if (document.datos.paginaDerFin.value == "0"){
      document.datos.paginader.value = parseInt(document.datos.paginader.value,10) + 1;
      actualizarSelectados();
   }
}

function antPagSel(){
   if (parseInt(document.datos.paginader.value,10) > 1){
      document.datos.paginader.value = parseInt(document.datos.paginader.value,10) - 1;
      actualizarSelectados();
   }
}

function actualizarTodo(){
   actualizar();
   actualizarSelectados();   
}

function cambioVer(){
  if (document.datos.tipovent[0].checked){
     document.datos.ventpagina.readOnly = false;
     document.datos.ventpagina.className = "habinp";	 
  }else{
     document.datos.ventpagina.readOnly = true;
     document.datos.ventpagina.className = "deshabinp";	   
  }
  actualizarTodo();
}

function Tecla(num){
  if (num==13) {
     if (document.datos.ventpagina.value != document.datos.ventpaginaant.value){
         if (!validanumero(document.datos.ventpagina, 4, 0)){
	        alert("El tamaño de ventana es invalido.");	
			document.datos.ventpagina.value = document.datos.ventpaginaant.value;
		 }else{
 	  	    if (parseInt(document.datos.ventpagina.value, 10) < 1){
		        alert("El tamaño de ventana debe ser mayor a cero.");	
				document.datos.ventpagina.value = document.datos.ventpaginaant.value;
			}else{
			    document.datos.ventpaginaant.value = document.datos.ventpagina.value;
	            document.datos.paginader.value = 1;			
	            document.datos.paginaizq.value = 1;						
				actualizarTodo();
			}
		 }
		 return false;
	 }	
  }
  return num;
}

function cambioTamPagina(){
     if (document.datos.ventpagina.value != document.datos.ventpaginaant.value){
         if (!validanumero(document.datos.ventpagina, 4, 0)){
	        alert("El tamaño de ventana es invalido.");	
			document.datos.ventpagina.value = document.datos.ventpaginaant.value;
		 }else{
 	  	    if (parseInt(document.datos.ventpagina.value, 10) < 1){
		        alert("El tamaño de ventana debe ser mayor a cero.");	
				document.datos.ventpagina.value = document.datos.ventpaginaant.value;
			}else{
			    document.datos.ventpaginaant.value = document.datos.ventpagina.value;
	            document.datos.paginader.value = 1;			
	            document.datos.paginaizq.value = 1;						
				actualizarTodo();			
			}
		 }
		 return false;
	 }	
}


function HabilitarDer(estado){
  if (estado){

     document.all.TodosAgr.href = "javascript:Todos(nselfil.registro.nselfil, selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);";
  	 document.all.UnoAgr.href   = "javascript:Uno(nselfil.registro.nselfil,selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);";
	 document.all.UnoQui.href   = "javascript:Uno(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);";
	 document.all.TodosQui.href = "javascript:Todos(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);";

  }else{

     nselfil.registro.selfil.readOnly = true;
     document.all.TodosAgr.href = "javascript:;";
  	 document.all.UnoAgr.href   = "javascript:;";
	 document.all.UnoQui.href   = "javascript:;";
	 document.all.TodosQui.href = "javascript:;";

  }
}

function HabilitarIzq(estado){
  if (estado){
  
     document.all.TodosAgr.href = "javascript:Todos(nselfil.registro.nselfil, selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);";
  	 document.all.UnoAgr.href   = "javascript:Uno(nselfil.registro.nselfil,selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);";
	 document.all.UnoQui.href   = "javascript:Uno(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);";
	 document.all.TodosQui.href = "javascript:Todos(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);";
  
  }else{

     nselfil.registro.nselfil.readOnly = true;
     document.all.TodosAgr.href = "javascript:;";
  	 document.all.UnoAgr.href   = "javascript:;";
	 document.all.UnoQui.href   = "javascript:;";
	 document.all.TodosQui.href = "javascript:;";

  }
}

function organizarTernros(){
     document.datos.accion.value = "U";
     document.datos.target = 'valida';
	 document.datos.action = 'gen_select_emp_v2_04.asp';
	 document.datos.submit();
}

function filtrosCargados(){
   if (document.datos.filtrosListo.value == "0"){
      setTimeout("filtrosCargados()",200);   
   }else{
      actualizar();
   }
}

function obtenerTerceros(){
  var arr = document.datos.seleccion.value.split(",");
  var arr2;
  var i;
  var lista1 = "0";
  
  for (i=1; i<arr.length; i++){
	  arr2 = arr[i].split('@');
	  lista1 = lista1 + ',' + arr2[0];
  }
  
  document.datos.seleccion.value = lista1;

}

</script>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<style>
.hidden2
{
	background : transparent;
	border : none;
	FONT-WEIGHT: bold;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%><%= l_titulo%></title>
</head>
<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">

<form name="datos" method="post" action="" target="">
<input type="Hidden" value="0" name="seleccion">
<input type="Hidden" value="<%= l_filtroIni%>" name="sqlfiltroemp">
<input type="Hidden" value="<%= l_filtroOnly%>" name="sqlfiltrofijo">
<input type="Hidden" value="" name="sqlfiltroizq">
<input type="Hidden" value="" name="sqlfiltroder">
<input type="Hidden" value="" name="sqlordenizq">
<input type="Hidden" value="" name="sqlordender">
<input type="Hidden" value="<%= l_selalto%>"  name="selalto">
<input type="Hidden" value="<%= l_selancho%>" name="selancho">
<input type="Hidden" value="0" name="paginaIzqFin">
<input type="Hidden" value="0" name="paginaDerFin">

<input type="Hidden" value="" name="accion">
<input type="Hidden" value="<%= l_vent_pagina%>" name="listanueva">
<input type="Hidden" value="<%= l_vent_pagina%>" name="ventpaginaant">
<input type="Hidden" value="0" name="filtrosListo">


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
         <iframe frameborder="1" name="ifrmfiltros" src="gen_select_emp_v2_01.asp?seltipnro=<%= l_seltipnro%>" width="98%" height="100" scroll=no></iframe><br>
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
    <table style="border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;" cellpadding="0" cellspacing="0" border="0">    
	  <tr>
  
    <td align="left" width="45%" valign="top" height="2%">
	<a class=sidebtnSHW href="javascript:filtroIz();">Filtro</a>
	<!--
	<a class=sidebtnSHW href="#" onClick="lado=1; MenuPopUp('elMenu1',event)" onMouseOut="lado=1; MenuPopDown('elMenu1')">Orden</a>
	-->
	<br><br>
	<b>No Seleccionados</b><br>
	Total:&nbsp;
	<input type="Text" name="totalizq" size="6" class="hidden" readonly>
	Total P&aacute;gina:&nbsp;
    <input type="Text" name="filtroizq" size="15" class="hidden" readonly>
	</td>
    <td align=left width="10%">
    </td>
    <td width="45%" align="left" valign="top" height="2%">
	<a class=sidebtnSHW href="javascript:filtroDe();">Filtro</a></a>
	<!--
	<a class=sidebtnSHW href="#" onClick="lado=0; MenuPopUp('elMenu1',event)" onMouseOut="lado=0; MenuPopDown('elMenu1')">Orden</a>
	-->
	<a class=sidebtnSHW href="javascript:agregarCriterio();">Agregar Criterio</a></a><br><br>	
	<b>Seleccionados</b><br>
    Total:&nbsp;
    <input type="Text" name="totalder" size="6" class="hidden" readonly>
	Total P&aacute;gina:&nbsp;
    <input type="Text" name="filtroder" size="15" class="hidden" readonly>
	</td>
</tr>
<tr>
  <td valign="top" align="center">
    <iframe align="center" name="nselfil" width="<%= l_selancho%>" height="<%= l_selalto%>" src="" scroll=no></iframe>   
	<br>
	<a href="javascript:antPagNoSel();"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Anterior" border="0"></a>	
	<input type="Text" name="paginaizq" size="3" class="hidden" value="1" readonly style="text-align: right; vertical-align: top;">
	<input type="Text" size="2" class="hidden" value="de" readonly style="text-align: center; vertical-align: top;">
	<input type="Text" name="totpaginaizq" size="3" class="hidden" value="1" readonly style="text-align: left; vertical-align: top;">	
	<a href="javascript:sigPagNoSel();"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Siguiente" border="0"></a>
  </td>
  <td valign="middle" align="center" width="10%" nowrap>
	<a name="TodosAgr" class=sidebtnSHW href="javascript:Todos(nselfil.registro.nselfil, selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);">>></a><br>
	<a name="UnoAgr"   class=sidebtnSHW href="javascript:Uno(nselfil.registro.nselfil,selfil.registro.selfil,document.datos.totalizq, document.datos.totalder);">></a><br>
	<a name="UnoQui"   class=sidebtnSHW href="javascript:Uno(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);"><</a></a><br>
	<a name="TodosQui" class=sidebtnSHW href="javascript:Todos(selfil.registro.selfil,nselfil.registro.nselfil,document.datos.totalder,document.datos.totalizq);"><<</a><br>
	<br><br>
	<table cellpadding="0" cellspacing="0" border="0" style="border-color:gray ; border-width: 1 ; border-style:solid ; margin-top: 5px;">
	  <tr>
	    <td nowrap>
			<input type="Radio" name="tipovent" value="0" checked onclick="javascript:cambioVer()"><b>Ver</b>&nbsp;<br>
			&nbsp;<input type="Text" name="ventpagina" value="<%= l_vent_pagina%>" size="4" maxlength=4"" onchange="javascript:cambioTamPagina();" onKeyPress="return Tecla(event.keyCode)"><br>
			<br>
			<input type="Radio" name="tipovent" value="1" onclick="javascript:cambioVer()"><b>Todos</b>	
		</td>
	  </tr>
	</table>
  </td>  
  <td valign="top" align="center">
    <iframe align="center" name="selfil" width="<%= l_selancho%>" height="<%= l_selalto%>" src=""></iframe>   
	<br>
	<a href="javascript:antPagSel();"><img align="absmiddle" src="/serviciolocal/shared/images/prev.jpg" alt="Anterior" border="0"></a>	
	<input type="Text" name="paginader" size="3" class="hidden" value="1" readonly style="text-align: right; vertical-align: top;">
	<input type="Text" size="2" class="hidden" value="de" readonly style="text-align: center; vertical-align: top;">
	<input type="Text" name="totpaginader" size="3" class="hidden" value="1" readonly style="text-align: left; vertical-align: top;">	
	<a href="javascript:sigPagSel();"><img align="absmiddle" src="/serviciolocal/shared/images/next.jpg" alt="Siguiente" border="0"></a>
  </td>
</tr>    

</table>
</td>
</tr>
<tr>
    <td align="right" class="th2" colspan="3" height="2%">
		<a name="botAceptar" class=sidebtnABM href="javascript:Aceptar('<%= l_srcDatos%>','<%= l_funtion%>','<%= l_ventana%>','<%= l_formulario%>')">Aceptar</a>
		<a class=sidebtnABM href="Javascript:cerrar();">Cancelar</a>
	</td>
</tr>
</table>
</form>

<iframe name="valida" width="0" height="0"></iframe>

<script>
  document.datos.seleccion.value = eval('<%= l_srcDatos & ".value"%>');
  if (document.datos.seleccion.value == ""){
      document.datos.seleccion.value = '0';
  }
   actualizarSelectados();

  document.datos.target = 'nselfil';
  document.datos.action = 'gen_select_emp_v2_02.asp';
  document.datos.submit();  
</script>

<SCRIPT SRC="/serviciolocal/shared/js/menu_op.js"></SCRIPT>

</html>

