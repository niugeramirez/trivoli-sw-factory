function Validaciones_locales_prov(){

	if ($("#nombreproveedor").val() == ""){
	//if (document.datos_02.nombre.value == ""){
		alert("Debe ingresar el Nombre del Proveedor.");
		document.datos_02.nombre.focus();
		return false;
	}

	return true;
}

function Submit_Formulario_prov() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'proveedores_con_06.asp'					//url_valid_06
						,'proveedores_con_03.asp'					//url_AM
						,'dialogAlert'							//id_dialogAlert
						,'datos_02_prov'								//id_form_datos
						,null //window.parent.ifrm.location			//location_reload
						,Validaciones_locales_prov					//funcion_Validaciones_locales
						,"ifrm"											//id_ifrm_form_datos
					);
} 