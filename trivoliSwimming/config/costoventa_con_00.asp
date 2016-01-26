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
Dim l_id

l_idVenta = request.querystring("id")
  %>

<html>
<head>

<title>Administracion de Costos de Ventas</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>
<script src="/trivoliSwimming/shared/js/fn_numeros.js"></script>

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_locales(){

	if (document.datos_02.idconceptoCompraVenta.value == "0"){
		alert("Debe ingresar un Concepto.");
		document.datos_02.idconceptoCompraVenta.focus();
		return false;
	}
	
	if (document.datos_02.cantidad.value == ""){
		alert("Debe ingresar una Cantidad.");
		document.datos_02.cantidad.focus();
		return false;
	}	
	
	if (document.datos_02.cantidad.value == "0"){
		alert("La Cantidad debe ser distinta de Cero.");
		document.datos_02.cantidad.focus();
		return false;
	}	
	
	document.datos_02.cantidad2.value = document.datos_02.cantidad.value.replace(",", ".");
	if (!validanumero(document.datos_02.cantidad2, 15, 4)){
		  alert("La Cantidad no es válida. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.cantidad.focus();
		  document.datos.cantidad.select();
		  return;
	}		
	
	if (document.datos_02.precio_unitario.value == ""){
		alert("Debe ingresar un Precio Unitario.");
		document.datos_02.precio_unitario.focus();
		return false;
	}		
	
	if (document.datos_02.precio_unitario.value == ""){
		alert("El Precio Unitario debe ser distinto de Cero.");
		document.datos_02.precio_unitario.focus();
		return false;
	}	
	
	document.datos_02.precio_unitario2.value = document.datos_02.precio_unitario.value.replace(",", ".");
	if (!validanumero(document.datos_02.precio_unitario2, 15, 4)){
		  alert("El Precio Unitario no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.precio_unitario.focus();
		  document.datos.precio_unitario.select();
		  return;
	}	
	
		
/*
	if (document.datos_02.idtemplatereserva.value == 0){
		alert("Debe ingresar el Modelo.");
		document.datos_02.idtemplatereserva.focus();
		return false;
	}

	if (document.datos_02.cantturnossimult.value == ""){
		alert("Debe ingresar la Cantidad de Turnos Simultaneos.");
		document.datos_02.cantturnossimult.focus();
		return false;
	}
	*/
	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'costoVenta_con_06.asp'					//url_valid_06
						,'costoVenta_con_03.asp'					//url_AM
						,'dialogAlert'							//id_dialogAlert
						,'datos_02'								//id_form_datos
						,window.parent.ifrm.location			//location_reload
						,Validaciones_locales					//funcion_Validaciones_locales
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete"				//id_dialogConfirmDelete
																	,"costoVenta_con_04.asp"				//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"costoVenta_con_06.asp"							//url_valid_06
														,"costoVenta_con_03.asp"							//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtro_00").val("");

	// Nombre
	if ($("#inpnombre").val() != 0){
		$("#filtro_00").val(" conceptosCompraVenta.descripcion like '*" + $("#inpnombre").val() + "*'");
	}		
    
	window.ifrm.location = 'costoVenta_con_01.asp?idventa=<%= l_idventa %>&asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'costoVenta_con_01.asp?idventa=<%= l_idventa %>';
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
                <input type="hidden" id="filtro_00" name="filtro_00" value="">
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
						<td><input  type="text" id="inpnombre" name="inpnombre" size="21" maxlength="21" value="" ></td>
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','costoVenta_con_02.asp?idventa=<%= l_idVenta %>&Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Cliente"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="costoVenta_con_01.asp?idventa=<%= l_idventa %>" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Costos de Ventas"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
