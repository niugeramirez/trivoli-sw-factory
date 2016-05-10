<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/inc/sec.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/const.inc"-->
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

  Dim l_rs
  Dim l_sql
  
  Dim l_alta
  Dim l_dni
  'Dim l_hc
  Dim l_fn_asign_pac
  Dim l_dnioblig
  Dim l_hcoblig  
  
  l_alta  = request("Alta")
  l_dni  = request("dni")
  'l_hc  = request("hc")
  l_fn_asign_pac = request("fn_asign_pac")  
  l_dnioblig  = request("dnioblig")
  l_hcoblig  = request("hcoblig")
  
  'response.write "l_fn_asign_pac "&l_fn_asign_pac
  'response.write "l_alta "&l_alta
  
  
%>
<html>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Buscar Proveedores</title>
<!--	VENTANAS MODALES        -->
<script>
function AsignarProveedor(id, nombre){

	<%= l_fn_asign_pac %>(id, nombre);
}
function devolver_proveedor_editado(id){

	<%= l_fn_asign_pac %>(	id,
							document.datos_02_prov.nombre.value 
							);
}

function Validaciones_locales_EditPro(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_prov();
}

function Submit_Formulario_EditPac() {
	Validar_Formulario(	"dialog_cont_EditPro" 										//id_dialog
						,"proveedores_con_06.asp"				//url_valid_06
						,"proveedores_con_03.asp"				//url_AM
						,"dialogAlert_BusqEdicPro"									//id_dialogAlert															
						,"datos_02_prov"										//id_form_datos							
						,null //window.parent.ifrm.location			//location_reload						
						,Validaciones_locales_EditPro					//funcion_Validaciones_locales						
						,"ifrm"											//id_ifrm_form_datos
						,devolver_paciente_editado //fn_post_AM	
					);
} 
$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_BusqEdicPro"									//id_dialogAlert
														);

								inicializar_dialogoABM(	"dialog_cont_EditPro" 										//id_dialog
														,"proveedores_con_06.asp"				//url_valid_06
														,"proveedores_con_03.asp"				//url_AM
														,"dialogAlert_BusqEdicPro"									//id_dialogAlert															
														,"datos_02_prov"										//id_form_datos															
														,null //window.parent.ifrm.location					//location_reload														
														,Validaciones_locales_EditPro							//funcion_Validaciones_locales	
														,"ifrm"											//id_ifrm_form_datos	
														,devolver_proveedor_editado //fn_post_AM														
														); 																
							});
</script>
<!--	FIN VENTANAS MODALES    -->
<script>

function Buscar(){
	var tieneotro;
	var estado;
	document.form_BusqPro.filtro.value = "";
	tieneotro = "no";
	estado = "si";



	// Nombre
	if (document.form_BusqPro.nombre.value != 0){
		if (tieneotro == "si"){
			document.form_BusqPro.filtro.value += " AND proveedores.nombre like '*" + document.form_BusqPro.nombre.value + "*'";
		}else{
			document.form_BusqPro.filtro.value += " proveedores.nombre like '*" + document.form_BusqPro.nombre.value + "*'";
		}
		tieneotro = "si";
	}		
				
				

	
	if (document.form_BusqPro.filtro.value.trim() == ""){
		alert("Debe ingresar el Filtro.");
		document.form_BusqPro.nombre.focus();
		return;
	}
	
	if (estado == "si"){
		window.ifrm_BusqPro.location = 'BuscarproveedoresV2_01.asp?asistente=0&filtro=' + document.form_BusqPro.filtro.value;
	}
}


function AltaProveedor(){
	abrirDialogo('dialog_cont_EditPro','proveedores_con_02.asp?Tipo=A&ventana=3&dni=<%= l_dni %>&hcoblig=<%= l_hcoblig %>',600,300);
}

function Limpiar(){
	
	window.form_BusqPro.location = 'BuscarproveedoresV2_01.asp?asistente=0&filtro=1=2' ;
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.form_BusqCli.nombre.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">	
		<tr>
			<td align="center" colspan="2">
				
					<form name="form_BusqPro" id="form_BusqPro" action="nada" target="nada">
						<input type="hidden" name="filtro" value="">
						<table border="0">
							<tr>
								<td align="right"><b>Nombre: </b></td>
								<td align="left" colspan="3" ><input  type="text" name="nombre" size="21" maxlength="21" value="" ></td>					

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
					<a href="Javascript:AltaProveedor();"><img src="../shared/images/Agregar_24.png" border="0" title="Alta Cliente"></a>	
				<% End If %>
			</td>			
		</tr>				
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
			<iframe scrolling="yes" name="ifrm_BusqPro" id="ifrm_BusqPro"  src="BuscarproveedoresV2_01.asp?asistente=0&filtro=1=1" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		

      </table>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->		
		<div id="dialogAlert_BusqEdicPro" title="Mensaje">				</div>			
		<div id="dialog_cont_EditPro" title="Editar Proveedores">		</div>	
		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>
</html>
