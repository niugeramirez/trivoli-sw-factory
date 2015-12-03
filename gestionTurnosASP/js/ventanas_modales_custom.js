/////

function Validar_Formulario(){

	if (!Validaciones_locales()){
		return;
		}
	else{ 
		$.post($("#url_valid_06").val(), $( "#datos" ).serialize(), 
				function(data) {     
									
									if(data=="OK") {
										valido();
									}
									else {
										abrirAlert("ERROR: " + data);
									}							
							});		
		}			
}

$(function () {

	$("#"+$("#id_dialogAlert").val()).dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {							
											$(this).dialog("close");
											}
			}
	});
	
	$("#"+$("#id_dialogConfirmDelete").val()).dialog({
		autoOpen: false,
		modal: true,
		buttons: {
					"Aceptar": function () {
											$(this).dialog("close");
											//$.post("obrassocialesV2_04.asp?cabnro=" + $("#cabnro_00").val()/*document.ifrm.datos.cabnro.value*/, {},
											$.post($("#url_baja").val(), {},											
													function(data) {     
																		
																		if(data=="OK") {
																			window.parent.ifrm.location.reload();
																		}
																		else {
																			$("#url_baja").val("0");
																			abrirAlert("ERROR: " + data);
																		}
																
																});												
																					
											},
					"Cerrar": function () {
											$(this).dialog("close");
											}											
			}
	});

	//$("#dialog").dialog({
	$("#"+$("#id_dialog").val()).dialog({
		autoOpen: false,
		modal: true,		
		buttons: {
					"Aceptar": function () {
											Validar_Formulario();											
											},
					"Cerrar": function () {
											$(this).dialog("close");
											$(this).empty();
											}
			}
	});
	

	
});
function abrirAlert(texto){
		$("#"+$("#id_dialogAlert").val()).text("");
		$("#"+$("#id_dialogAlert").val()).dialog("option", "width", 600);
		$("#"+$("#id_dialogAlert").val()).dialog("option", "height", 300);
		$("#"+$("#id_dialogAlert").val()).dialog("option", "resizable", false);
		$("#"+$("#id_dialogAlert").val()).dialog("open");
		$("#"+$("#id_dialogAlert").val()).html(texto);	
}	
function abrirConfirmDelete(texto){
		$("#"+$("#id_dialogConfirmDelete").val()).text("");
		$("#"+$("#id_dialogConfirmDelete").val()).dialog("option", "width", 600);
		$("#"+$("#id_dialogConfirmDelete").val()).dialog("option", "height", 300);
		$("#"+$("#id_dialogConfirmDelete").val()).dialog("option", "resizable", false);
		$("#"+$("#id_dialogConfirmDelete").val()).dialog("open");
		$("#"+$("#id_dialogConfirmDelete").val()).html(texto);	
}		
function abrirDialogo(url){
		$("#"+$("#id_dialog").val()).text("");
		$("#"+$("#id_dialog").val()).dialog("option", "width", $("#width_dialog").val() );
		$("#"+$("#id_dialog").val()).dialog("option", "height", $("#height_dialog").val() );
		$("#"+$("#id_dialog").val()).dialog("option", "resizable", false);
		$("#"+$("#id_dialog").val()).dialog("open");
		//$("#dialog").load(url);	
		$.ajax({	url: url,
					cache: false,
					dataType: "html",
					success: function(data) {$("#"+$("#id_dialog").val()).html(data);}
				});		
}
function abrirDialogoVerif(url) 
{
  if (ifrm.jsSelRow == null)
    abrirAlert("Debe seleccionar un registro.")
  else	   
    abrirDialogo(url) 
}

function eliminarRegistroAJAX(obj)
{
	if (obj.datos.cabnro.value == 0)
		{
		abrirAlert("Debe seleccionar un registro para realizar la operaci&oacute;n.");
		}
	else
		{
			//$("#cabnro_00").val(obj.datos.cabnro.value);
			$("#url_baja").val($("#url_base_baja").val()+"cabnro="+obj.datos.cabnro.value); 
			abrirConfirmDelete("&iquest;Desea eliminar el registro seleccionado?");
		}
}

function valido(){
	
	$.post($("#url_AM").val(), $( "#datos" ).serialize(),
			function(data) {     
								if(data=="OK") {
									$("#"+$("#id_dialog").val()).dialog("close"); 
									$("#"+$("#id_dialog").val()).empty();									
									window.parent.ifrm.location.reload();
								}
								else {
									abrirAlert("ERROR: " + data);
								}
						
						});
	//document.datos.submit();
}
//////