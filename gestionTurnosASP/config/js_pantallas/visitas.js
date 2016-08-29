function Validaciones_locales_AVST(){

	if (document.datosAVST.pacienteid.value == "0"){
		alert("Debe ingresar el Paciente.");
		document.datosAVST.pacienteid.focus();
		return;
	}

	if (document.datosAVST.practicaid.value == "0"){
		alert("Debe ingresar la Practica.");
		document.datosAVST.practicaid.focus();
		return;
	}

	document.datosAVST.precio2.value = document.datosAVST.precio.value.replace(",", ".");
	if (!validanumero(document.datosAVST.precio2, 15, 4)){
			  alert("El Precio no es válido. Se permite hasta 15 enteros y 4 decimales.");	
			  document.datosAVST.precio.focus();
			  document.datosAVST.precio.select();
			  return;
	}

	//Chequeo que existan los datos de pago que en las actualizaciones de practicas no s visualizan estos campos
	if (document.datosAVST.idmediodepago){
		if (document.datosAVST.mediodepagoos.value == document.datosAVST.idmediodepago.value)  {
			if (Trim(document.datosAVST.idobrasocial.value) == "0"){
				alert("Debe ingresar la Obra Social.");
				document.datosAVST.idobrasocial.focus();
				return;
			}
		}			

		if (document.datosAVST.importe.value == ""){
			if	(document.datosAVST.idmediodepago.value != "0") {
				alert("Debe ingresar un Importe mayor o igual a 0.");
				document.datosAVST.importe.focus();
				return;
			}
		}
	

		document.datosAVST.importe2.value = document.datosAVST.importe.value.replace(",", ".");

		if (!validanumero(document.datosAVST.importe2, 15, 4)){
				  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
				  document.datosAVST.importe.focus();
				  document.datosAVST.importe.select();
				  return;
		}	

		if (document.datosAVST.importe.value != 0)  {
			if (Trim(document.datosAVST.idmediodepago.value) == "0"){
				alert("Debe ingresar el Medio de Pago.");
				document.datosAVST.idmediodepago.focus();
				return;
			}
		}

		if (document.datosAVST.idmediodepago.value != "0")  {
			if (Trim(document.datosAVST.importe.value) == "0"){
				alert("Debe ingresar el Importe.");
				document.datosAVST.importe.focus();
				return;
			}
		}
	}

	return true;
}

function Submit_Formulario_visit() {
	Validar_Formulario(	'dialogVisit'								//id_dialog
						,'visitasV2_con_06.asp'					//url_valid_06
						,'visitasV2_con_03.asp'					//url_AM
						,'dialogAlertVisitas'							//id_dialogAlert
						,'datosAVST'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_locales_AVST					//funcion_Validaciones_locales
						,"ifrm_visit"											//id_ifrm_form_datos
					);				
} 

function actualizarprecio(p_precio){	
	document.datosAVST.precio.value = p_precio;
	
	// Si el medio de Pago es Obra social, copio el precio al importe
	//Solo lo copio si existe el campo medio de pago ya que en las actualizaciones de practicas realizadas no existe
	if (document.datosAVST.idmediodepago) {
		if (document.datosAVST.idmediodepago.value == document.datosAVST.mediodepagoos.value ) { 
			document.datosAVST.importe.value = p_precio;
		} 
		else document.datosAVST.importe.value = 0;	
	}
}	



function calcularprecio(){
	//document.valida.location = "agregarpractica_con_06.asp?idos=" + document.datosAVST.osid.value + "&practicaid="+ document.datosAVST.practicaid.value ;	
	$.post("query_precio_practica_JSON.asp?idos=" + document.datosAVST.osid.value + "&practicaid="+ document.datosAVST.practicaid.value, 
		function(data) {
							if (EsJsonString(data)) {							
								if($.parseJSON(data)[0].resultado=="OK") {
									actualizarprecio($.parseJSON(data)[0].precio);									
								}
								else {
									abrirAlert('dialogAlertVisitas',"ERROR: " +data);
								}
							}								
							else {
								abrirAlert('dialogAlertVisitas',"ERROR: " +data);
							}								
					});	
}

function volver_AsignarPaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os){

	document.datosAVST.pacienteid.value = id;
	document.datosAVST.apellido.value = apellido;
	document.datosAVST.nombre.value = nombre;
	document.datosAVST.nrohistoriaclinica.value = nrohistoriaclinica;
	document.datosAVST.dni.value = dni;
	document.datosAVST.domicilio.value = domicilio;
	document.datosAVST.tel.value = tel;
	document.datosAVST.osid.value = osid;
	document.datosAVST.os.value = os;
	document.datosAVST.idobrasocial.value = osid;
	
	// lo dejo dividido asi por si mas adelante deshabilitamos algun control
	if (osid == document.datosAVST.osparticular.value){
		document.datosAVST.idobrasocial.value = 0;
		document.datosAVST.idmediodepago.value = 0;
	}
    else {
		document.datosAVST.idobrasocial.value = osid;
		document.datosAVST.idmediodepago.value = document.datosAVST.mediodepagoos.value;
	};	
	$("#dialog_cont_BusqPac").dialog("close");
}


function ctrolmetodopago(){
	//chequeo que exista el campo medio de pago, ya que en la modificacion de practicas estos campos no se muestran
	if (document.datosAVST.idmediodepago) {
		if (document.datosAVST.mediodepagoos.value == document.datosAVST.idmediodepago.value) {		
				document.datosAVST.idobrasocial.disabled = false;							
			}
			else {
				document.datosAVST.idobrasocial.disabled = true;							
				document.datosAVST.idobrasocial.value = 0;	
			}			
	}
}
////////////////////////////////////////////FUNCIONES PARA EL ALTA DE VISITAS CON TURNO/////////////////////////////////////////////////////////
function Habilitar(obj,  turno, permitir){	


	if (permitir=="N") {
		alert('El Paciente seleccionado no tiene DNI o Nro de Historia Clinica cargado. Ir a la opcion Pacientes para completar esta informacion');
		obj.checked="";
		return;
	}
		
	if (obj.checked==false) {
	//alert('es falso');
	document.datos_02_AVCT.cabnro.value = document.datos_02_AVCT.cabnro.value.replace(','+turno, '');
	}
	else {
	//alert('es verdadero ');
	document.datos_02_AVCT.cabnro.value = document.datos_02_AVCT.cabnro.value + "," + turno ;
	document.datos_02_AVCT.cabnro2.value = document.datos_02_AVCT.cabnro2.value.replace(','+turno, '');
	};
	
	if (!obj.checked) return 
    elem=document.getElementsByName(obj.name); 
    for(i=0;i<elem.length;i++)  
        elem[i].checked=false; 
    obj.checked=true; 		
}

function Habilitar2(obj,  turno, permitir){	

	
	if (obj.checked==false) {
	//alert('es falso');
	document.datos_02_AVCT.cabnro2.value = document.datos_02_AVCT.cabnro2.value.replace(','+turno, '');
	}
	else {
	//alert('es verdadero ');
	document.datos_02_AVCT.cabnro2.value = document.datos_02_AVCT.cabnro2.value + "," + turno ;
	document.datos_02_AVCT.cabnro.value = document.datos_02_AVCT.cabnro.value.replace(','+turno, '');
	};
	
	if (!obj.checked) return 
    elem=document.getElementsByName(obj.name); 
    for(i=0;i<elem.length;i++)  
        elem[i].checked=false; 
    obj.checked=true; 		
}

function Validaciones_locales_AVconT(){

	if ((Trim(document.datos_02_AVCT.cabnro.value) == "0") && (Trim(document.datos_02_AVCT.cabnro2.value) == "0")){
		alert("Debe seleccionar alguna Opcion.");		
		return;
	}	

	return true;
}