<% Option Explicit %>
<% 
'Archivo: historia_clinica_resum_con_00.asp
'Descripción: Resumen de Historias Clinicas para Impresion
'Autor : Trivoli
'Fecha: 02/02/2016

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql
  %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Resumen de Historias Clinicas para Impresion</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
///////////////////////////////////////EDICION PACIENTES   ///////////////////////////////////////
function Validaciones_locales_EditPac_HCR(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_EditPac_02()
}
function devolver_paciente_editado_HCR(id,obj){
	volver_AsignarPaciente(	document.datos_02_EditPac.id.value, 
							document.datos_02_EditPac.apellido.value, 
							document.datos_02_EditPac.nombre.value, 
							obj.nrohistoriaclinica, //document.datos_02_EditPac.nrohistoriaclinica.value, 
							document.datos_02_EditPac.dni.value, 
							document.datos_02_EditPac.domicilio.value, 
							document.datos_02_EditPac.tel.value, 
							document.datos_02_EditPac.osid.value, 
							document.datos_02_EditPac.os.value
							);
}
function Editar_Paciente(){ 

	if (document.datos_02.idclientepaciente.value == 0){
		alert("Debe ingresar el Paciente.");
		document.datos_02.idclientepaciente.focus();
		return;
	}; 
		 
	abrirDialogo('dialogHCR_cont_EditPac','EditarpacientesV2_02.asp?Tipo=M&ventana=3&dnioblig=S&hcoblig=S&cabnro='+document.datos_02.idclientepaciente.value,600,300);
}
///////////////////////////////////////FIN EDICION PACIENTES///////////////////////////////////////
function Validaciones_locales(){

	if (document.datos_02.idrecursoreservable.value == "0"){
		alert("Debe ingresar un Medico.");
		document.datos_02.idrecursoreservable.focus();
		return false;
	}

	if (document.datos_02.detalle.value == ""){
		alert("Debe ingresar el Detalle.");
		document.datos_02.detalle.focus();
		return false;
	}

	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'historia_clinica_resum_con_06.asp'		//url_valid_06
						,'historia_clinica_resum_con_03.asp'		//url_AM
						,'dialogAlert'							//id_dialogAlert
						,'datos_02'								//id_form_datos
						,window.parent.ifrm.location			//location_reload
						,Validaciones_locales					//funcion_Validaciones_locales
						,"ifrm"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete"				//id_dialogConfirmDelete
																	,"historia_clinica_resum_con_04.asp"	//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"historia_clinica_resum_con_06.asp"				//url_valid_06
														,"historia_clinica_resum_con_03.asp"				//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos														
														); 
								inicializar_dialogoContenedor(	"dialog_cont_BusqPac" 										//id_dialog
																); 		
								inicializar_dialogoABM(	"dialogHCR_cont_EditPac" 										//id_dialog
														,"EditarpacientesV2_06.asp"				//url_valid_06
														,"EditarpacientesV2_03_JSON.asp"				//url_AM
														,"dialogAlert"									//id_dialogAlert															
														,"datos_02_EditPac"										//id_form_datos															
														,null //window.parent.ifrm.location					//location_reload														
														,Validaciones_locales_EditPac_HCR							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos	
														,devolver_paciente_editado_HCR //fn_post_AM														
														); 		
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtro_00").val("");

	// Apellido
	if ($("#inpapellido").val() != 0){
		$("#filtro_00").val(" clientespacientes.apellido like '*" + $("#inpapellido").val() + "*'");
	}		
    
	window.ifrm.location = 'historia_clinica_resum_con_01.asp?asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'historia_clinica_resum_con_01.asp';
}


function volver_AsignarPaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os){
	document.datos_02.idclientepaciente.value = id;
	document.datos_02.apellidoclientepaciente.value = apellido;
	document.datos_02.nombreclientepaciente.value = nombre;
	
	$("#dialog_cont_BusqPac").dialog("close");
}

function BuscarPaciente(){	
	abrirDialogo('dialog_cont_BusqPac','BuscarpacientesV2_00.asp?Tipo=A&Alta=S&fn_asign_pac=volver_AsignarPaciente&dnioblig=N&hcoblig=N',900,250);
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
                            Resumen de Historias Clinicas para Impresion
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
					    <td  align="right" ><b>Paciente: </b></td>
						<td><input  type="text" id="inpapellido" name="inpapellido" size="21" maxlength="21" value="" ></td>
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="../shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="../shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','historia_clinica_resum_con_02.asp?Tipo=A',850,600)"><img  src="../shared/images/Agregar_24.png" border="0" title="Agregar Historia Clinica"></a>    


                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="historia_clinica_resum_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Historias Clinicas"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>	
		<div id="dialog_cont_BusqPac" title="Buscar Pacientes">		</div>	
		<div id="dialogHCR_cont_EditPac" title="Editar Pacientes">		</div>	
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
