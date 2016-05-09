<% Option Explicit %>
<% 
'Archivo: compras_con_00.asp
'Descripción: Administración de Compras
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

<title>Administracion de Compras</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>


<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_locales(){

	if (document.datos_02.idproveedor.value == "0"){
		alert("Debe ingresar el Proveedor.");
		document.datos_02.idproveedor.focus();
		return false;
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
						,'compras_con_06.asp'					//url_valid_06
						,'compras_con_03.asp'					//url_AM
						,'dialogAlert'							//id_dialogAlert
						,'datos_02'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_locales					//funcion_Validaciones_locales
						,"ifrm"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete"				//id_dialogConfirmDelete
																	,"compras_con_04.asp"				//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"compras_con_06.asp"							//url_valid_06
														,"compras_con_03.asp"							//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos
														); 

								inicializar_dialogoContenedor(	"dialog_cont_DC" 										//id_dialog
																); 
								inicializar_dialogoContenedor(	"dialog_cont_BusqPro" 										//id_dialog
																); 																	
								inicializar_dialogoContenedor(	"dialog_cont_CMC" 										//id_dialog
																); 
																
								//esta linea la agrego solo para refrescar cuando se cierra el dialogo contenedor, se podría parametrizar de modo de recibir
								//la funcion como parametro que se debe ejecutar al
								$( "#dialog_cont_DC" ).dialog({
									close: function () {$(this).empty(); Buscar();}
								});	
								$( "#dialog_cont_CMC" ).dialog({
									close: function () {$(this).empty(); Buscar();}
								});
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtro_00").val("");

	// Nombre
	if ($("#inpnombre").val() != 0){
		//Eugenio 02/05/2016 en lugar de enviar * para los comodines envio **, con esto el el 01 reemplazo el ** por %
		//Esto es porque el otro filtro implica una operacion matematica de multiplicacion, con lo que se modifica el * y hace un calculo erroneo
		$("#filtro_00").val(" proveedores.nombre like '**" + $("#inpnombre").val() + "**'");
	}		
    
 
	if ($("#filt_con_saldo_comp").is(':checked'))
	{
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+ " ( " 
								+ " isnull((SELECT  "
								+ "     SUM(detalleCompras.cantidad * detalleCompras.precio_unitario) "
								+ " FROM detalleCompras "
								+ "   WHERE detalleCompras.idcompra = compras.id),0)  "
								+ " -  "
								+ " isnull((SELECT "
								+ "     SUM(cajaMovimientos.monto) "
								+ "   FROM cajaMovimientos "
								+ "   WHERE cajaMovimientos.idcompraOrigen = compras.id),0)  "
								+ " <>0								 "
								+ " ) " 
							);		
	}

	window.ifrm.location = 'compras_con_01.asp?asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'compras_con_01.asp';
}

function volver_AsignarProveedor(id,  nombre){

	document.datos_02.idproveedor.value = id;
	document.datos_02.proveedor.value = nombre;
	
	$("#dialog_cont_BusqPro").dialog("close");
}

function BuscarProveedor(){	
	
	abrirDialogo('dialog_cont_BusqPro','BuscarproveedoresV2_00.asp?Tipo=A&Alta=N&fn_asign_pac=volver_AsignarProveedor&dnioblig=N&hcoblig=N',900,250);
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
                            Administracion de Compras
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
					    <td align="right"><b>Proveedor: </b></td>
						<td><input  type="text" id="inpnombre" name="inpnombre" size="60" maxlength="21" value="" ></td>
					    <td>
							<b>Con saldo:</b><input type="checkbox" id="filt_con_saldo_comp" name="filt_con_saldo_comp">
						</td>
                        <td align="left">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','compras_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Compra"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="compras_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Compras"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>		

		<div id="dialog_cont_DC" title="Detalle de Compras">		</div>		
		<div id="dialog_cont_CMC" title="Pagos">		</div>			
	    <div id="dialog_cont_BusqPro" title="Buscar Proveedores">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
