<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: obrassociales_00.asp
'Descripción: ABM de Obras Sociales
'Autor : RAUL CHINESTRA
'Fecha: 19/04/2005
'Modificado: 

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
  l_etiquetas = "Codigo:;Descripción:"
  l_Campos    = "balcod;baldes"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "Código:;Descripción:"
  l_CamposOr  = "balcod;baldes"

%>
<html>
<head>

<title>Obras Sociales</title>


<!--<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.1/themes/base/jquery-ui.css" />-->
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<!--<script src="http://code.jquery.com/jquery-1.9.1.js"></script>-->
<script src="../js/jquery.min.js"></script>
<!--<script src="http://code.jquery.com/ui/1.10.1/jquery-ui.js"></script>-->
<script src="../js/jquery-ui.js"></script>

<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script>
/////
function Validar_Formulario(){

	if (Trim(document.datos.descripcion.value) == ""){
		abrirAlert("Debe ingresar la Descripci&oacute;n de la Obra Social.");
		document.datos.descripcion.focus();
		return;
		}
	else{
		$.post("obrassocialesV2_06.asp", {tipo: document.datos.tipo.value, id: document.datos.id.value, descripcion: document.datos.descripcion.value},   
				function(data) {     
									
									if(data=="OK") {
										valido();
									}
									else {
										abrirAlert("ERROR: " + data);
									}
							
							});		
		}
	
		//valido();
			
}
//////
$(function () {

	$("#dialogAlert").dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {							
											$(this).dialog("close");
											}
			}
	});
	
	$("#dialogConfirmDelete").dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {
											$(this).dialog("close");
											//$.post("obrassocialesV2_04.asp?cabnro=" + $("#cabnro_00").val()/*document.ifrm.datos.cabnro.value*/, {},
											$.post($("#url_baja").val(), {},											
													function(data) {     
																		
																		if(data=="OK") {
																			window.parent.ifrm.location.reload();
																		}
																		else {
																			$("#url_baja").val("0");
																			abrirAlert("ERROR: " + data);
																		}
																
																});												
																					
											},
					"Cerrar": function () {
											$(this).dialog("close");
											}											
			}
	});

	$("#dialog").dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {
											Validar_Formulario();											
											},
					"Cerrar": function () {
											$(this).dialog("close");
											}
			}
	});
	

	
});
function abrirAlert(texto){
		$("#dialogAlert").text("");
		$("#dialogAlert").dialog("option", "width", 600);
		$("#dialogAlert").dialog("option", "height", 300);
		$("#dialogAlert").dialog("option", "resizable", false);
		$("#dialogAlert").dialog("open");
		$("#dialogAlert").html(texto);	
}	
function abrirConfirmDelete(texto){
		$("#dialogConfirmDelete").text("");
		$("#dialogConfirmDelete").dialog("option", "width", 600);
		$("#dialogConfirmDelete").dialog("option", "height", 300);
		$("#dialogConfirmDelete").dialog("option", "resizable", false);
		$("#dialogConfirmDelete").dialog("open");
		$("#dialogConfirmDelete").html(texto);	
}		
function abrirDialogo(url){
		$("#dialog").text("");
		$("#dialog").dialog("option", "width", 600);
		$("#dialog").dialog("option", "height", 300);
		$("#dialog").dialog("option", "resizable", false);
		$("#dialog").dialog("open");
		//$("#dialog").load(url);	
		$.ajax({	url: url,
					cache: false,
					dataType: "html",
					success: function(data) {$("#dialog").html(data);}
				});		
}
function abrirDialogoVerif(url) 
{
  if (ifrm.jsSelRow == null)
    abrirAlert("Debe seleccionar un registro.")
  else	   
    abrirDialogo(url) 
}

function eliminarRegistroAJAX(obj)
{
	if (obj.datos.cabnro.value == 0)
		{
		abrirAlert("Debe seleccionar un registro para realizar la operaci&oacute;n.");
		}
	else
		{
			//$("#cabnro_00").val(obj.datos.cabnro.value);
			$("#url_baja").val("obrassocialesV2_04.asp?cabnro="+obj.datos.cabnro.value); 
			abrirConfirmDelete("&iquest;Desea eliminar el registro seleccionado?");
		}
}

function valido(){
	
	$.post("obrassocialesV2_03.asp", {tipo: document.datos.tipo.value, id: document.datos.id.value, descripcion: document.datos.descripcion.value},   
			function(data) {     
								if(data=="OK") {
									$("#dialog").dialog("close");  
									window.parent.ifrm.location.reload();
								}
								else {
									abrirAlert("ERROR: " + data);
								}
						
						});
	//document.datos.submit();
}
</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra"></td>
          <td nowrap align="right" class="barra">
          		  
		  <a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('obrassocialesV2_02.asp?Tipo=A')"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta"></a>
		  &nbsp;
          
		  <!-- Este bloque no va mas porque se llevan las bajas y modificaciones a la grilla
		  <a href="Javascript:abrirDialogoVerif('obrassocialesV2_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value);"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar"></a>
		  &nbsp;
          		 
		  <a href="Javascript:eliminarRegistroAJAX(document.ifrm);"><img src="/turnos/shared/images/Eliminar_16.png" border="0" title="Baja"></a>								  
		  &nbsp;&nbsp;
  		  
          <a href="Javascript:abrirVentanaVerif('listadeprecios_con_00.asp?id=' + document.ifrm.datos.cabnro.value,'',520,200);"><img src="/turnos/shared/images/Ecommerce-Price-Tag-icon_24.png" border="0" title="Lista de Precios"></a>								  
			-->
		  </td>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe  id="ifrm" name="ifrm" src="obrassocialesV2_01.asp" width="100%" height="100%"></iframe> 
		  <input type="hidden" id="url_baja" value="0">
	      </td>
        </tr>
		
      </table>
	  
		<div id="dialog" title="Obras Sociales">

		</div>	  
		
		<div id="dialogAlert" title="Mensaje">

		</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">

		</div>			
</body>
</html>
