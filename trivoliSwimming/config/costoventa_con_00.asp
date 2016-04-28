<% Option Explicit %>
<% 
'Archivo:costoVenta_con_00.asp
'Descripción: Administración de Detalle de Costo de Ventas
'Autor : Trivoli
'Fecha: 31/05/2015

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql

Dim l_idVenta
Dim l_p_carga_js
Dim l_id

l_idVenta = request.querystring("id")
l_p_carga_js = request.querystring("p_carga_js")
'response.write  "l_p_carga_js "&l_p_carga_js&"</br>"
  %>

<html>
<head>

<title>Administracion de Costos de Ventas</title>
<% if l_p_carga_js = "S" or isnull(l_p_carga_js) or l_p_carga_js = "" then%>
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>


<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>
<% end if %>
<script src="/trivoliSwimming/shared/js/fn_numeros.js"></script>
<script>
function Validaciones_locales_cv(){

	if (document.datos_02_cv.idconceptoCompraVenta.value == "0"){
		alert("Debe ingresar un Concepto.");
		document.datos_02_cv.idconceptoCompraVenta.focus();
		return false;
	}
	
	if (document.datos_02_cv.cantidad.value == ""){
		alert("Debe ingresar una Cantidad.");
		document.datos_02_cv.cantidad.focus();
		return false;
	}	
	
	if (document.datos_02_cv.cantidad.value == "0"){
		alert("La Cantidad debe ser distinta de Cero.");
		document.datos_02_cv.cantidad.focus();
		return false;
	}	
	
	document.datos_02_cv.cantidad2.value = document.datos_02_cv.cantidad.value.replace(",", ".");
	if (!validanumero(document.datos_02_cv.cantidad2, 15, 4)){
		  alert("La Cantidad no es válida. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.cantidad.focus();
		  document.datos.cantidad.select();
		  return;
	}		
	
	if (document.datos_02_cv.precio_unitario.value == ""){
		alert("Debe ingresar un Precio Unitario.");
		document.datos_02_cv.precio_unitario.focus();
		return false;
	}		
	
	if (document.datos_02_cv.precio_unitario.value == ""){
		alert("El Precio Unitario debe ser distinto de Cero.");
		document.datos_02_cv.precio_unitario.focus();
		return false;
	}	
	
	document.datos_02_cv.precio_unitario2.value = document.datos_02_cv.precio_unitario.value.replace(",", ".");
	if (!validanumero(document.datos_02_cv.precio_unitario2, 15, 4)){
		  alert("El Precio Unitario no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.precio_unitario.focus();
		  document.datos.precio_unitario.select();
		  return;
	}	
	
	return true;
}

function Submit_Formulario_cv() {
	Validar_Formulario(	'dialog_cv'								//id_dialog
						,'costoVenta_con_06.asp'					//url_valid_06
						,'costoVenta_con_03.asp'					//url_AM
						,'dialogAlert_cv'							//id_dialogAlert
						,'datos_02_cv'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_locales_cv					//funcion_Validaciones_locales
						,"ifrm_cv"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_cv"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete_cv"				//id_dialogConfirmDelete
																	,"costoVenta_con_04.asp"				//url_baja
																	,"dialogAlert_cv"						//id_dialogAlert
																	,"detalle_01_cv"						//id_form_datos
																	,"ifrm_cv"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog_cv" 										//id_dialog
														,"costoVenta_con_06.asp"							//url_valid_06
														,"costoVenta_con_03.asp"							//url_AM
														,"dialogAlert_cv"									//id_dialogAlert	
														,"datos_02_cv"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_cv							//funcion_Validaciones_locales	
														,"ifrm_cv"											//id_ifrm_form_datos														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar_cv(){

	$("#filtro_00_cv").val("");

	// Nombre
	if ($("#inpnombre_cv").val() != 0){
		$("#filtro_00_cv").val(" conceptosCompraVenta.descripcion like '*" + $("#inpnombre_cv").val() + "*'");
	}		
    
	window.ifrm_cv.location = 'costoVenta_con_01.asp?idventa=<%= l_idventa %>&asistente=0&filtro=' + $("#filtro_00_cv").val();
}

function Limpiar_cv(){
	window.ifrm_cv.location = 'costoVenta_con_01.asp?idventa=<%= l_idventa %>';
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
                            Administracion de Costos de Ventas
                        </td>
                    </tr>
                </table>
    		</td>
        </tr>
        <tr>
			<td>
                <input type="hidden" id="filtro_00_cv" name="filtro_00_cv" value="">
				<table border="0" width="100%">
                    <colgroup>
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                    </colgroup>
                    <tbody>
				    <tr>
					    <td><b>Concepto: </b></td>
						<td><input  type="text" id="inpnombre_cv" name="inpnombre_cv" size="21" maxlength="21" value="" ></td>
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar_cv();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar_cv();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog_cv','costoVenta_con_02.asp?idventa=<%= l_idVenta %>&Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Cliente"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm_cv" name="ifrm_cv" src="costoVenta_con_01.asp?idventa=<%= l_idventa %>" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog_cv" title="Costos de Ventas"> 			</div>	  
				
		<div id="dialogAlert_cv" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete_cv" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar_cv();
</script>
</html>
