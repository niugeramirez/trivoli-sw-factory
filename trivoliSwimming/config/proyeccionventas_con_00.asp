<% Option Explicit %>
<% 
'Archivo: proyeccionventas_con_00.asp
'Descripción: Administración de proyeccionventas
'Autor : Trivoli
'Fecha: 31/05/2015

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql
  %>

<html>
<head>

<title>Proyeccion de ventas</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>
<script src="/trivoliSwimming/shared/js/fn_numeros.js"></script>
<!-- Comienzo Datepicker -->
 
<script>
$(function () {
/*$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});*/

		
$( "#filt_fecha_desde" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#filt_fecha_hasta" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->


<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_localesPV(){

	if (document.datospv_02.fecha_desde.value == ""){
		alert("Debe ingresar Fecha Desde.");
		document.datospv_02.fecha_desde.focus();
		return false;
	}

	if (document.datospv_02.fecha_hasta.value == ""){
		alert("Debe ingresar Fecha Hasta.");
		document.datospv_02.fecha_desde.focus();
		return false;
	}
	
	if (document.datospv_02.fecha_hasta.value < document.datospv_02.fecha_desde.value ){
		alert("La Fecha Hasta debe ser menor a la Fecha Desde.");
		document.datospv_02.fecha_desde.focus();
		return false;
	}	
	
	if (document.datospv_02.cantidadproyectada.value == ""){
		alert("Debe ingresar una Cantidad.");
		document.datospv_02.cantidadproyectada.focus();
		return false;
	}	
	
	if (document.datospv_02.cantidadproyectada.value == "0"){
		alert("La Cantidad debe ser distinta de Cero.");
		document.datospv_02.cantidadproyectada.focus();
		return false;
	}	
	
	document.datospv_02.cantidadproyectada.value = document.datospv_02.cantidadproyectada.value.replace(",", ".");
	if (!validanumero(document.datospv_02.cantidadproyectada, 15, 4)){
		  alert("La Cantidad no es válida. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datospv_02.cantidadproyectada.focus();
		  document.datospv_02.cantidadproyectada.select();
		  return;
	}	
	
	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialogPV'								//id_dialog
						,'proyeccionventas_con_06.asp'					//url_valid_06
						,'proyeccionventas_con_03.asp'					//url_AM
						,'dialogAlertPV'							//id_dialogAlert
						,'datospv_02'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_localesPV					//funcion_Validaciones_locales
						,"ifrmpv"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlertPV"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDeletePV"				//id_dialogConfirmDelete
																	,"proyeccionventas_con_04.asp"				//url_baja
																	,"dialogAlertPV"						//id_dialogAlert
																	,"detallePV_01"						//id_form_datos
																	,"ifrmpv"								//id_ifrm_form_datos
																	,null //window.parent.ifrmpv.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialogPV" 										//id_dialog
														,"proyeccionventas_con_06.asp"							//url_valid_06
														,"proyeccionventas_con_03.asp"							//url_AM
														,"dialogAlertPV"									//id_dialogAlert	
														,"datospv_02"										//id_form_datos		
														,null //window.parent.ifrmpv.location					//location_reload
														,Validaciones_localesPV							//funcion_Validaciones_locales	
														,"ifrmpv"											//id_ifrm_form_datos														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtropv_00").val("");

	// fecha desde
	if ($("#filt_fecha_desde").val() != 0){
		$("#filtropv_00").val(" CONVERT(VARCHAR(10), proyeccionventas.fecha_desde, 111)  >= " + cambiafechaYYYYMMDD($("#filt_fecha_desde").val(),true,1) + "" );		
	}	
	// fecha desde
	if ($("#filt_fecha_hasta").val() != 0){
		if ($("#filt_fecha_hasta").val() !="") {
			$("#filtropv_00").val( $("#filtropv_00").val() + " AND ");
		}
		$("#filtropv_00").val( $("#filtropv_00").val() + " CONVERT(VARCHAR(10), proyeccionventas.fecha_hasta, 111)  <= " + cambiafechaYYYYMMDD($("#filt_fecha_hasta").val(),true,1) + "" );		
	}		    
	window.ifrmpv.location = 'proyeccionventas_con_01.asp?asistente=0&filtro=' + $("#filtropv_00").val();
}

function Limpiar(){
	window.ifrmpv.location = 'proyeccionventas_con_01.asp';
}
</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
    <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr>
            <td align="left">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td class="title">
                            Proyeccion de ventas
                        </td>
                    </tr>
                </table>
    		</td>
        </tr>
        <tr>
			<td>
                <input type="hidden" id="filtropv_00" name="filtropv_00" value="">
				<table border="0" width="100%">
                    <colgroup>
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                    </colgroup>
                    <tbody>
				    <tr>
					    <td><b>Fecha Desde: </b>
							<input  type="text" id="filt_fecha_desde" name="filt_fecha_desde" size="21" maxlength="21" value="" >						
						</td>
					    <td><b>Fecha Hasta: </b>
							<input  type="text" id="filt_fecha_hasta" name="filt_fecha_hasta" size="21" maxlength="21" value="" >						
						</td>						
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialogPV','proyeccionventas_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Banco"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrmpv" name="ifrmpv" src="proyeccionventas_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialogPV" title="Proyeccion"> 			</div>	  
				
		<div id="dialogAlertPV" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDeletePV" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
