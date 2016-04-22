<% Option Explicit %>
<% 
'Archivo: cajamovimientos_con_00.asp
'Descripci�n: Administraci�n de cajamovimientos
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

<title>Administracion de Caja Movimientos</title>

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

		
$( "#filt_fechadesde_cm" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#filt_fechahasta_cm" ).datepicker({
	showOn: "button",
	buttonImage: "/trivoliSwimming/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_locales(){

	if (document.datos_02.fecha.value == ""){
		alert("Debe ingresar la fecha del movimiento.");
		document.datos_02.fecha.focus();
		return false;
	}	
	if (document.datos_02.tipoes.value == ""){
		alert("Debe ingresar el Tipo.");
		document.datos_02.tipoes.focus();
		return false;
	}	
	
	if (document.datos_02.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02.idtipomovimiento.focus();
		return false;
	}	
	
	if (document.datos_02.detalle.value == ""){
		alert("Debe ingresar el Detalle.");
		document.datos_02.detalle.focus();
		return false;
	}		
	
	if (document.datos_02.idunidadnegocio.value == "0"){
		alert("Debe ingresar la Unidad de Negocio.");
		document.datos_02.idunidadnegocio.focus();
		return false;
	}	
	
	if (document.datos_02.idmediopago.value == "0"){
		alert("Debe ingresar el Medio de Pago.");
		document.datos_02.idmediopago.focus();
		return false;
	}	
	
	if (document.datos_02.mediodepagocheque.value == document.datos_02.idmediopago.value){
		if (document.datos_02.idcheque.value == "0"){
			alert("Debe ingresar el Cheque.");
			document.datos_02.idcheque.focus();
			return false;
		}	

	}		
	
	if (document.datos_02.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02.idtipomovimiento.focus();
		return false;
	}				

	if (document.datos_02.monto.value == ""){
		alert("Debe ingresar un Monto.");
		document.datos_02.monto.focus();
		return false;
	}		
	
	if (document.datos_02.monto.value == "0"){
		alert("El Monto debe ser distinto de Cero.");
		document.datos_02.monto.focus();
		return false;
	}		
	
	document.datos_02.monto2.value = document.datos_02.monto.value.replace(",", ".");
	if (!validanumero(document.datos_02.monto2, 15, 4)){
		  alert("El Monto no es v�lido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.monto.focus();
		  document.datos.monto.select();
		  return;
	}		
	
	if (document.datos_02.idresponsable.value == "0"){
		alert("Debe ingresar el Responsable.");
		document.datos_02.idresponsable.focus();
		return false;
	}	


	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'cajamovimientos_con_06.asp'					//url_valid_06
						,'cajamovimientos_con_03.asp'					//url_AM
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
																	,"cajamovimientos_con_04.asp"				//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"cajamovimientos_con_06.asp"							//url_valid_06
														,"cajamovimientos_con_03.asp"							//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtro_00").val("");

	// Nombre
	if ($("#filt_detalle_cm").val() != 0){
		$("#filtro_00").val(" cajamovimientos.detalle like '*" + $("#filt_detalle_cm").val() + "*'");
	}		
    
	//Fecha desde
	if ($("#filt_fechadesde_cm").val() != 0){
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+ " cajamovimientos.fecha  >= " + cambiafechaYYYYMMDD($("#filt_fechadesde_cm").val(),true,1)
							);		
	}	

	//Fecha hasta
	if ($("#filt_fechahasta_cm").val() != 0){
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+ " cajamovimientos.fecha  <= " + cambiafechaYYYYMMDD($("#filt_fechahasta_cm").val(),true,1)
							);		
	}	
	
	//Numero de cheque 
	if ($("#filt_nrocheque_cm").val() != 0){
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+" cheques.numero like '*" + $("#filt_nrocheque_cm").val() + "*'"
							);		
	}	
	
	//banco 
	if ($("#filt_banco_cm").val() != 0){
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+" bancos.nombre_banco like '*" + $("#filt_banco_cm").val() + "*'"
							);		
	}
	
	//Cl�iente/Proveedor 
	if ($("#filt_cliprov_cm").val() != 0){
		if ($("#filtro_00").val() != 0){
			$("#filtro_00").val( $("#filtro_00").val() + " and ");
		}
		$("#filtro_00").val(
								$("#filtro_00").val() 
								+"( proveedores.nombre like '*" + $("#filt_cliprov_cm").val() + "*'"
								+" or clientes.nombre like '*" + $("#filt_cliprov_cm").val() + "*')"
							);		
	}
	
	window.ifrm.location = 'cajamovimientos_con_01.asp?asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'cajamovimientos_con_01.asp';
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
                            Administracion de Caja
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
					    <td>
							<b>Cliente: </b><input  type="text" id="filt_cliprov_cm" name="filt_cliprov_cm" size="21" maxlength="21" value="" >
						</td>
					    <td>
							<b>Fecha: </b><input id="filt_fechadesde_cm" type="text" name="filt_fechadesde_cm" size="10" maxlength="10" value="" >							
						</td>
						<td>
							<b>Nro Cheque: </b><input id="filt_nrocheque_cm" type="text" name="filt_nrocheque_cm" size="10" maxlength="10" value="" >							
						</td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','cajamovimientos_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Cliente"></a>    
                        </td>
                    </tr>
					<tr>
						<td>
							<b>Detalle: </b><input  type="text" id="filt_detalle_cm" name="filt_detalle_cm" size="21" maxlength="21" value="" >
						</td>
					    <td>					
							<b>Hasta: </b><input id="filt_fechahasta_cm" type="text" name="filt_fechahasta_cm" size="10" maxlength="10" value="" >	
						</td>						
						<td>
							<b>Banco&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp: </b><input id="filt_banco_cm" type="text" name="filt_banco_cm" size="10" maxlength="10" value="" >	
						</td>
						<td></td>
					</tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="cajamovimientos_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Cajas"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
