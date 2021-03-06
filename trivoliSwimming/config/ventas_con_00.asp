<% Option Explicit %>
<% 
'Archivo: ventas_con_00.asp
'Descripción: Administración de Ventas
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

<title>Administracion de Ventas</title>

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
///////////////////////////////////////EDICION Clientes   ///////////////////////////////////////
function Validaciones_locales_EditCli_HCR(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_EditCli_02()
}
function devolver_cliente_editado_HCR(){
	volver_AsignarCliente(	document.datos_02_EditCli.id.value, 
							//document.datos_02_EditCli.apellido.value, 
							document.datos_02_EditCli.nombre.value 
							//document.datos_02_EditCli.nrohistoriaclinica.value, 
							//document.datos_02_EditCli.dni.value, 
							//document.datos_02_EditCli.domicilio.value, 
							//document.datos_02_EditCli.tel.value, 
							//document.datos_02_EditCli.osid.value, 
							//document.datos_02_EditCli.os.value
							);
}
function Editar_Cliente(){ 

	if (document.datos_02.idcliente2.value == 0){
		alert("Debe ingresar el Cliente.");
		document.datos_02.idcliente2.focus();
		return;
	}; 
		
	abrirDialogo('dialogHCR_cont_EditCli','EditarclientesV2_02.asp?Tipo=M&ventana=3&dnioblig=N&hcoblig=N&cabnro='+document.datos_02.idcliente2.value,600,300);
}
///////////////////////////////////////FIN EDICION PACIENTES///////////////////////////////////////
function Validaciones_locales(){

	if (document.datos_02.fecha.value == ""){
		alert("Debe ingresar una Fecha.");
		document.datos_02.idcliente2.focus();
		return false;
	}

	if ((document.datos_02.fecha.value == "")&&(!validarfecha(document.datos_02.fecha))){
		 document.datos_02.fecha.focus();
		 return;
	}

	if (document.datos_02.idcliente2.value == "0"){
		alert("Debe ingresar el Cliente.");
		document.datos_02.idcliente2.focus();
		return false;
	}

	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'ventas_con_06.asp'					//url_valid_06
						,'ventas_con_03.asp'					//url_AM
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
																	,"ventas_con_04.asp"				//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"ventas_con_06.asp"							//url_valid_06
														,"ventas_con_03.asp"							//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos														
														); 

								inicializar_dialogoContenedor(	"dialog_cont_DV" 										//id_dialog
																); 

								inicializar_dialogoContenedor(	"dialog_cont_CV" 										//id_dialog
																); 																

								inicializar_dialogoContenedor(	"dialog_cont_CM" 										//id_dialog
																); 	
																
								inicializar_dialogoContenedor(	"dialog_cont_BusqCli" 										//id_dialog
																); 		

								inicializar_dialogoContenedor(	"dialog_cont_BusqCli" 										//id_dialog
																); 																	
																
								inicializar_dialogoABM(	"dialogHCR_cont_EditCli" 										//id_dialog
														,"EditarclientesV2_06.asp"				//url_valid_06
														,"EditarclientesV2_03_JSON.asp"				//url_AM
														,"dialogAlert"									//id_dialogAlert															
														,"datos_02_EditCli"										//id_form_datos															
														,null //window.parent.ifrm.location					//location_reload														
														,Validaciones_locales_EditCli_HCR							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos	
														,devolver_cliente_editado_HCR //fn_post_AM															
														); 															
																
								//esta linea la agrego solo para refrescar cuando se cierra el dialogo contenedor, se podría parametrizar de modo de recibir
								//la funcion como parametro que se debe ejecutar al
								$( "#dialog_cont_DV" ).dialog({
									close: function () {$(this).empty(); Buscar();}
								});		
								$( "#dialog_cont_CV" ).dialog({
									close: function () {$(this).empty(); Buscar();}
								});	
								$( "#dialog_cont_CM" ).dialog({
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
		$("#filtro_00").val(" clientes.nombre like '**" + $("#inpnombre").val() + "**'");
	}		

	//saldo
	if ($("#filt_con_saldo_vent").is(':checked'))
	{
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+ " ( " 
								+ " isnull((SELECT  "
								+ "     SUM(detalleVentas.cantidad * detalleVentas.precio_unitario) "
								+ " FROM detalleVentas "
								+ "   WHERE detalleVentas.idVenta = ventas.id),0)  "
								+ " -  "
								+ "   isnull((SELECT "
								+ "     SUM(cajaMovimientos.monto) "
								+ "   FROM cajaMovimientos "
								+ "   WHERE cajaMovimientos.idventaOrigen = ventas.id),0)  "
								+ " <>0								 "
								+ " ) " 	
							);		
	}	
   
	window.ifrm.location = 'ventas_con_01.asp?asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'ventas_con_01.asp';
}

function volver_AsignarCliente(id,  nombre){

	document.datos_02.idcliente2.value = id;
	document.datos_02.cliente.value = nombre;
	
	$("#dialog_cont_BusqCli").dialog("close");
}


function BuscarCliente(){	
	abrirDialogo('dialog_cont_BusqCli','BuscarclientesV2_00.asp?Tipo=A&Alta=S&fn_asign_pac=volver_AsignarCliente&dnioblig=N&hcoblig=N',900,250);
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
                            Administracion de Ventas
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
					    <td align="right"><b>Cliente: </b></td>
						<td><input  type="text" id="inpnombre" name="inpnombre" size="60" maxlength="21" value="" ></td>
					    <td><b>Con saldo:</b><input type="checkbox" id="filt_con_saldo_vent" name="filt_con_saldo_vent">
						</td>
                        <td align="left">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','ventas_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Venta"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="ventas_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Ventas"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>	

		<div id="dialog_cont_DV" title="Detalle de Ventas">		</div>		
		
		<div id="dialog_cont_CV" title="Costos Ventas">		</div>		
		
		<div id="dialog_cont_CM" title="Pagos">		</div>	
		
		<div id="dialog_cont_BusqCli" title="Buscar Clientes">		</div>			
		
		<div id="dialogHCR_cont_EditCli" title="Editar Cliente">		</div>			
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
