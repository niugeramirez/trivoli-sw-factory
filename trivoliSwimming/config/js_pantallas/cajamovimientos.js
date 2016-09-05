function volver_AsignarCheque(id, fecha,  numero, banco, importe ){

	document.datos_02_mc.cheque_nom.value = numero + " - " + banco + " - " + fecha;
	document.datos_02_mc.idcheque.value = id;
	document.datos_02_mc.monto.value = importe;
	
	$("#dialog_cont_BusqCheque").dialog("close");
}

function devolver_cheque_editado(){
	volver_AsignarCheque(	document.datos_02_cheq.id.value,
							document.datos_02_cheq.fecha_emision.value, 
							document.datos_02_cheq.numero.value, 
							$( "#idbanco option:selected" ).text().trim(),
							document.datos_02_cheq.importe.value  
							);							
}	

function Validaciones_locales_mc(){

	if (document.datos_02_mc.fecha.value == ""){
		alert("Debe ingresar la fecha del movimiento.");
		document.datos_02_mc.fecha.focus();
		return false;
	}	
	if (document.datos_02_mc.tipoes.value == ""){
		alert("Debe ingresar el Tipo.");
		document.datos_02_mc.tipoes.focus();
		return false;
	}	
	
	if (document.datos_02_mc.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02_mc.idtipomovimiento.focus();
		return false;
	}	
	
	/*if (document.datos_02_mc.detalle.value == ""){
		alert("Debe ingresar el Detalle.");
		document.datos_02_mc.detalle.focus();
		return false;
	}*/		
	
	/*if (document.datos_02_mc.idunidadnegocio.value == "0"){
		alert("Debe ingresar la Unidad de Negocio.");
		document.datos_02_mc.idunidadnegocio.focus();
		return false;
	}	
	*/
	
	if (document.datos_02_mc.idmediopago.value == "0"){
		alert("Debe ingresar el Medio de Pago.");
		document.datos_02_mc.idmediopago.focus();
		return false;
	}	
	
	if (document.datos_02_mc.mediodepagocheque.value == document.datos_02_mc.idmediopago.value){
		if (document.datos_02_mc.idcheque.value == "0"){
			alert("Debe ingresar el Cheque.");
			document.datos_02_mc.idcheque.focus();
			return false;
		}	

	}		
	
	if (document.datos_02_mc.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02_mc.idtipomovimiento.focus();
		return false;
	}				

	if (document.datos_02_mc.monto.value == ""){
		alert("Debe ingresar un Monto.");
		document.datos_02_mc.monto.focus();
		return false;
	}		
	
	if (document.datos_02_mc.monto.value == "0"){
		alert("El Monto debe ser distinto de Cero.");
		document.datos_02_mc.monto.focus();
		return false;
	}		
	
	document.datos_02_mc.monto2.value = document.datos_02_mc.monto.value.replace(",", ".");
	if (!validanumero(document.datos_02_mc.monto2, 15, 4)){
		  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.monto.focus();
		  document.datos.monto.select();
		  return;
	}		
	
	if (document.datos_02_mc.idresponsable.value == "0"){
		alert("Debe ingresar el Responsable.");
		document.datos_02_mc.idresponsable.focus();
		return false;
	}	


	return true;
}

function Submit_Formulario_mc() {
	Validar_Formulario(	'dialog_mc'								//id_dialog
						,'cajamovimientos_con_06.asp'					//url_valid_06
						,'cajamovimientos_con_03.asp'					//url_AM
						,'dialogAlert_mc'							//id_dialogAlert
						,'datos_02_mc'								//id_form_datos
						,null //window.parent.ifrm_mc.location			//location_reload
						,Validaciones_locales_mc					//funcion_Validaciones_locales
						,"ifrm_mc"											//id_ifrm_form_datos
					);
} 







function BuscarCheque(){	
	
	abrirDialogo('dialog_cont_BusqCheque','BuscarChequeV2_00.asp?Tipo=A&Alta=S&fn_asign_pac=volver_AsignarCheque',900,250);		
		
}

function volver_AsignarCompraOrigen(id, fecha,  nombre){

	document.datos_02_mc.compraorigen.value = nombre + " - " + fecha ;
	document.datos_02_mc.idcompraorigen.value = id;
	
	$("#dialog_cont_BusqCompraOrigen").dialog("close");
}

function volver_AsignarVentaOrigen(id, fecha,  nombre){

	document.datos_02_mc.ventaorigen.value = nombre + " - " + fecha ;
	document.datos_02_mc.idventaorigen.value = id;
	
	$("#dialog_cont_BusqVentaOrigen").dialog("close");
}


function Editar_Cheque(){ 

	if (document.datos_02_mc.idcheque.value == 0){
		alert("Debe ingresar el Cheque.");
		document.datos_02_,c.idcheque.focus();
		return;
	}; 
		
	abrirDialogo('dialog_cont_EditCheq_CM','cheques_con_02.asp?Tipo=M&cabnro='+document.datos_02_mc.idcheque.value,600,300);
}

// /CONTROLES DEL 02
function ctrolmediodepago(){

	if (document.datos_02_mc.mediodepagocheque.value == document.datos_02_mc.idmediopago.value) {			
			//document.datos_02_mc.idcheque.disabled = false;	
			mostrar('cheque');						
		}
		else {
		
			//document.datos_02_mc.idcheque.disabled = true;	
			cerrar('cheque');								
	
		}	

}

function ctrolcheque(){

document.valida.location = "importecheque_con_00.asp?id=" + document.datos_02_mc.idcheque.value ;	

}

function actualizarimporte(p_importe){	
	document.datos_02_mc.monto.value = p_importe;
}


function ctroltipomovimiento(){

	document.valida.location = "flagtipomovimiento_con_00.asp?id=" + document.datos_02_mc.idtipomovimiento.value ;	

}

function actualizarflag (p_flagcompra, p_flagventa){	
	
		if (p_flagcompra == -1 ) {			
			document.datos_02_mc.idcompraorigen.disabled = false;		
			mostrar('compraorigen');					
		}
		else {			
			document.datos_02_mc.idcompraorigen.disabled = true;	
			document.datos_02_mc.idcompraorigen.value = 0;
			cerrar('compraorigen');					
		};

		if (p_flagventa == -1 ) {			
			document.datos_02_mc.idventaorigen.disabled = false;	
			mostrar('ventaorigen');						
		}
		else {			
			document.datos_02_mc.idventaorigen.disabled = true;	
			document.datos_02_mc.idventaorigen.value = 0;
			cerrar('ventaorigen');	
			
		};	
	
}
