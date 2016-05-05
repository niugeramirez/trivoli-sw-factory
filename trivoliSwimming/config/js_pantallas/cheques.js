function Validaciones_locales_cheq(){

	
	if (document.datos_02_cheq.numero.value == ""){
		alert("Debe ingresar un Numero.");
		document.datos_02_cheq.numero.focus();
		return false;
	}	
	if (document.datos_02_cheq.fecha_emision.value == ""){
		alert("Debe ingresar Fecha Emision.");
		document.datos_02_cheq.fecha_emision.focus();
		return false;
	}	
	if (document.datos_02_cheq.fecha_vencimiento.value == ""){
		alert("Debe ingresar Fecha Vencimiento.");
		document.datos_02_cheq.fecha_vencimiento.focus();
		return false;
	}		
	if (document.datos_02_cheq.numero.value == ""){
		alert("Debe ingresar un Numero.");
		document.datos_02_cheq.numero.focus();
		return false;
	}		
	if (document.datos_02_cheq.idbanco.value == "0"){
		alert("Debe ingresar un Banco.");
		document.datos_02_cheq.idbanco.focus();
		return false;
	}

	if (document.datos_02_cheq.importe.value == ""){
		alert("Debe ingresar un Importe.");
		document.datos_02_cheq.importe.focus();
		return false;
	}		
	
	if (document.datos_02_cheq.importe.value == "0"){
		alert("El Importe debe ser distinto de Cero.");
		document.datos_02_cheq.importe.focus();
		return false;
	}		
	
	document.datos_02_cheq.importe2.value = document.datos_02_cheq.importe.value.replace(",", ".");
	if (!validanumero(document.datos_02_cheq.importe2, 15, 4)){
		  alert("El Importe no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.importe.focus();
		  document.datos.importe.select();
		  return;
	}		
	
	if (document.datos_02_cheq.emisor.value == "0"){
		alert("Debe ingresar un Emisor.");
		document.datos_02_cheq.emisor.focus();
		return false;
	}	

	return true;
}

function Submit_Formulario_cheq() {
	Validar_Formulario(	'dialog_cheq'								//id_dialog
						,'cheques_con_06.asp'					//url_valid_06
						,'cheques_con_03.asp'					//url_AM
						,'dialogAlert_cheq'							//id_dialogAlert
						,'datos_02_cheq'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_locales_cheq					//funcion_Validaciones_locales
						,"ifrm_cheq"											//id_ifrm_form_datos
					);
} 