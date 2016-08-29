<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

  Dim l_rs
  Dim l_sql
  
  Dim l_alta
  'Dim l_dni
  'Dim l_hc
  Dim l_fn_asign_pac
  Dim l_dnioblig
  Dim l_hcoblig  
  
  l_alta  = request("Alta")
  'l_dni  = request("dni")
  'l_hc  = request("hc")
  l_fn_asign_pac = request("fn_asign_pac")  
  l_dnioblig  = request("dnioblig")
  l_hcoblig  = request("hcoblig")

'response.write "l_dnioblig "&l_dnioblig
'response.write "l_hcoblig "&l_hcoblig  
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Buscar Pacientes</title>
<!--	VENTANAS MODALES        -->
<script>
function AsignarPaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os){
	<%= l_fn_asign_pac %>(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os);
}
function devolver_paciente_editado(id,obj){
	<%= l_fn_asign_pac %>(	obj.id,//document.datos_02_EditPac.id.value, 
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

function Validaciones_locales_EditPac(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_EditPac_02()
}

function Submit_Formulario_EditPac() {
	Validar_Formulario(	"dialog_cont_EditPac" 										//id_dialog
						,"EditarpacientesV2_06.asp"				//url_valid_06
						,"EditarpacientesV2_03_JSON.asp"				//url_AM
						,"dialogAlert_BusqEdicPac"									//id_dialogAlert															
						,"datos_02_EditPac"										//id_form_datos							
						,null //window.parent.ifrm.location			//location_reload						
						,Validaciones_locales_EditPac					//funcion_Validaciones_locales						
						,"ifrm"											//id_ifrm_form_datos
						,devolver_paciente_editado //fn_post_AM	
					);
} 
$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_BusqEdicPac"									//id_dialogAlert
														);

								inicializar_dialogoABM(	"dialog_cont_EditPac" 										//id_dialog
														,"EditarpacientesV2_06.asp"				//url_valid_06
														,"EditarpacientesV2_03_JSON.asp"				//url_AM
														,"dialogAlert_BusqEdicPac"									//id_dialogAlert															
														,"datos_02_EditPac"										//id_form_datos															
														,null //window.parent.ifrm.location					//location_reload														
														,Validaciones_locales_EditPac							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos	
														,devolver_paciente_editado //fn_post_AM														
														); 																
							});
</script>
<!--	FIN VENTANAS MODALES    -->
<script>

function Buscar(){
	var tieneotro;
	var estado;
	document.form_BusqPac.filtro.value = "";
	tieneotro = "no";
	estado = "si";


	// Apellido
	if (document.form_BusqPac.legape.value != 0){
		if (tieneotro == "si"){
			document.form_BusqPac.filtro.value += " AND clientespacientes.apellido like '*" + document.form_BusqPac.legape.value + "*'";
		}else{
			document.form_BusqPac.filtro.value += " clientespacientes.apellido like '*" + document.form_BusqPac.legape.value + "*'";
		}
		tieneotro = "si";
	}		
	// Nombre
	if (document.form_BusqPac.legnom.value != 0){
		if (tieneotro == "si"){
			document.form_BusqPac.filtro.value += " AND clientespacientes.nombre like '*" + document.form_BusqPac.legnom.value + "*'";
		}else{
			document.form_BusqPac.filtro.value += " clientespacientes.nombre like '*" + document.form_BusqPac.legnom.value + "*'";
		}
		tieneotro = "si";
	}		
	// Nro. Historia Clinica
	if (document.form_BusqPac.nrohistoriaclinica.value != 0){
		if (tieneotro == "si"){
			document.form_BusqPac.filtro.value += " AND clientespacientes.nrohistoriaclinica like '*" + document.form_BusqPac.nrohistoriaclinica.value + "*'";
		}else{
			document.form_BusqPac.filtro.value += " clientespacientes.nrohistoriaclinica like '*" + document.form_BusqPac.nrohistoriaclinica.value + "*'";
		}
		tieneotro = "si";
	}				
	// DNI
	if (document.form_BusqPac.legdni.value != 0){
		if (tieneotro == "si"){
			document.form_BusqPac.filtro.value += " AND clientespacientes.dni like '*" + document.form_BusqPac.legdni.value + "*'";
		}else{
			document.form_BusqPac.filtro.value += " clientespacientes.dni like '*" + document.form_BusqPac.legdni.value + "*'";
		}
		tieneotro = "si";
	}					
				

	
	if (document.form_BusqPac.filtro.value.trim() == ""){
		alert("Debe ingresar el Filtro.");
		document.form_BusqPac.legape.focus();
		return;
	}
	
	if (estado == "si"){
		window.ifrm_BusqPac.location = 'BuscarpacientesV2_01.asp?asistente=0&filtro=' + document.form_BusqPac.filtro.value;
	}
}


function AltaPaciente(){

	abrirDialogo('dialog_cont_EditPac','EditarpacientesV2_02.asp?Tipo=A&ventana=3&dnioblig=<%= l_dnioblig %>&hcoblig=<%= l_hcoblig %>',600,300);
}

function Limpiar(){
	
	window.ifrm_BusqPac.location = 'BuscarpacientesV2_01.asp?asistente=0&filtro=1=2' ;
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.form_BusqPac.legape.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">	
		<tr>
			<td align="center" colspan="2">
				
					<form name="form_BusqPac" id="form_BusqPac" action="nada" target="nada">
						<input type="hidden" name="filtro" value="">
						<table border="0">
							<tr>
								<td align="right"><b>Apellido: </b></td>
								<td><input  type="text" name="legape" size="21" maxlength="21" value="" ></td>
								<td align="right"><b>Nombre: </b></td>
								<td><input  type="text" name="legnom" size="21" maxlength="21" value="" ></td>					

							</tr>
							<tr>
								<td align="right"><b>D.N.I.: </b></td>
								<td><input  type="text" name="legdni" size="21" maxlength="21" value="" >							</td>		
								<td align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
								<td><input type="text" name="nrohistoriaclinica" size="21" value="">							</td>											
							</tr>								
						</table>
					</form>		

			</td>
			<td align="center">
				<a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="../shared/images/Buscar_24.png" border="0" title="Buscar">
			</td>	
			<td align="center">	
				<a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="../shared/images/Limpiar_24.png" border="0" title="Limpiar">  
			</td>
			<td align="center">				
				<% if l_alta = "S" then %>
					<a href="Javascript:AltaPaciente();"><img src="../shared/images/Agregar_24.png" border="0" title="Alta Paciente"></a>	
				<% End If %>
			</td>			
		</tr>				
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
			<iframe scrolling="yes" name="ifrm_BusqPac" id="ifrm_BusqPac"  src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		

      </table>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->		
		<div id="dialogAlert_BusqEdicPac" title="Mensaje">				</div>			
		<div id="dialog_cont_EditPac" title="Editar Pacientes">		</div>	
		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>
</html>
