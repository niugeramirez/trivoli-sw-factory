/////

function Validar_Formulario(id_dialog,url_valid_06,url_AM,id_dialogAlert,id_form_datos,location_reload,funcion_Validaciones_locales,id_ifrm_form_datos,fn_post_AM){

	if (!funcion_Validaciones_locales()){
		return;
		}
	else{ 
		$.post(url_valid_06, $( "#"+id_form_datos).serialize(), 
				function(data) {     
									
									if(data=="OK") {
										valido(id_dialog,url_AM,id_dialogAlert,id_form_datos,location_reload,id_ifrm_form_datos,fn_post_AM);										
									}
									else {
										abrirAlert(id_dialogAlert,"ERROR: " + data);
									}							
							});		
		}			
}

function inicializar_dialogAlert(id_dialogAlert) {

	$("#"+id_dialogAlert).dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {							
											$(this).dialog("close");
											$(this).empty();
											}
			}
	});	
};

function inicializar_dialogConfirmDelete(id_dialogConfirmDelete,url_baja,id_dialogAlert,id_form_datos,id_ifrm_form_datos,location_reload) {
	
	$("#"+id_dialogConfirmDelete).dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {
											$(this).dialog("close");
											$(this).empty();											
											$.post(url_baja, $( "#"+id_form_datos,$("#"+id_ifrm_form_datos).contents()).serialize(),	
													function(data) {     
																		
																		if(data=="OK") {																			
																			$("#"+id_ifrm_form_datos).get(0).contentWindow.location.reload();//location_reload.reload();																		
																		}
																		else {
																			abrirAlert(id_dialogAlert,"ERROR: " + data);
																		}
																
																});												
																					
											},
					"Cerrar": function () {
											$(this).dialog("close");
											$(this).empty();
											}											
			}
	});
	
};
function inicializar_dialogoABM(id_dialog,url_valid_06,url_AM,id_dialogAlert,id_form_datos,location_reload,funcion_Validaciones_locales,id_ifrm_form_datos,fn_post_AM) {
		
	$("#"+id_dialog).dialog({		
		autoOpen: false,
		modal: true,		
		buttons: {
					"Aceptar": function () {
											Validar_Formulario(id_dialog,url_valid_06,url_AM,id_dialogAlert,id_form_datos,location_reload,funcion_Validaciones_locales,id_ifrm_form_datos,fn_post_AM);											
											},
					"Cerrar": function () {
											$(this).dialog("close");
											$(this).empty();
											}
			}
	});
	
};

function inicializar_dialogoContenedor(id_dialog) {
		
	$("#"+id_dialog).dialog({		
		autoOpen: false,
		modal: true,	
		close: function () {$(this).empty();}
	});
	
};

function abrirAlert(id_dialogAlert,texto){
		$("#"+id_dialogAlert).text("");
		$("#"+id_dialogAlert).dialog("option", "width", 600);
		$("#"+id_dialogAlert).dialog("option", "height", 300);
		$("#"+id_dialogAlert).dialog("option", "resizable", false);
		$("#"+id_dialogAlert).dialog("open");
		$("#"+id_dialogAlert).html(texto);	
}	
function abrirConfirmDelete(id_dialogConfirmDelete,texto){
		$("#"+id_dialogConfirmDelete).text("");
		$("#"+id_dialogConfirmDelete).dialog("option", "width", 600);
		$("#"+id_dialogConfirmDelete).dialog("option", "height", 300);
		$("#"+id_dialogConfirmDelete).dialog("option", "resizable", false);
		$("#"+id_dialogConfirmDelete).dialog("open");
		$("#"+id_dialogConfirmDelete).html(texto);	
}		


function abrirDialogo(id_dialog,url,width_dialog,height_dialog){
		$("#"+id_dialog).text("");		
		$("#"+id_dialog).dialog("option", "width", width_dialog);
		$("#"+id_dialog).dialog("option", "height", height_dialog);		
		$("#"+id_dialog).dialog("option", "resizable", false);
		$("#"+id_dialog).dialog("open");
		//$("#dialog").load(url);	
		$.ajax({	url: url,
					cache: false,
					dataType: "html",
					success: function(data) {$("#"+id_dialog).html(data);}
				});		
}
/*
//Este procedimiento no se utiliza, por eso lo comento. Si se comeinza a utilizar hay que parametrizarlo como al resto.
function abrirDialogoVerif(url) 
{
  if (ifrm.jsSelRow == null)
    abrirAlert("Debe seleccionar un registro.")
  else	   
    abrirDialogo(url) 
}
*/
function eliminarRegistroAJAX(obj_id,id_dialogAlert,id_dialogConfirmDelete)
{
	if (obj_id.value == 0)
		{
		abrirAlert(id_dialogAlert,"Debe seleccionar un registro para realizar la operaci&oacute;n.");
		}
	else
		{			 
			abrirConfirmDelete(id_dialogConfirmDelete,"&iquest;Desea eliminar el registro seleccionado?");
		}
}

function EsJsonString(str) {
    try {
        JSON.parse(str);
    } catch (e) {
        return false;
    }
    return true;
}

function valido(id_dialog,url_AM,id_dialogAlert,id_form_datos,location_reload,id_ifrm_form_datos,fn_post_AM){
	
	$.post(url_AM, $( "#"+id_form_datos ).serialize(),
			function(data) {     
								if(data=="OK") {
									$("#"+id_dialog).dialog("close"); 									
									$("#"+id_ifrm_form_datos).get(0).contentWindow.location.reload();//location_reload.reload();
									if (fn_post_AM) {
										fn_post_AM();
									}
									$("#"+id_dialog).empty();
								}
								else if (EsJsonString(data)) {
										if ($.parseJSON(data)[0].resultado=="OK"){
											$("#"+id_dialog).dialog("close"); 									
											$("#"+id_ifrm_form_datos).get(0).contentWindow.location.reload();//location_reload.reload();
											if (fn_post_AM) { fn_post_AM($.parseJSON(data)[0].id);}
											$("#"+id_dialog).empty();
										}
										else {
											abrirAlert(id_dialogAlert,"ERROR: " + data);
										}
								}
								else{
									abrirAlert(id_dialogAlert,"ERROR: " + data);
								}
						
						});
}


//////

////Seleccion de filas//////
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro,obj_id){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	//obj.datos.cabnro.value = cabnro;
	obj_id.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;	
	
}
////FIN Seleccion de filas//////

