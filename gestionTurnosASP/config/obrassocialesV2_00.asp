<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: obrassocialesV2_00.asp
'Descripción: ABM de Obras Sociales version modal
'Autor : Eugenio Ramirez
'Fecha: 31/08/2015


%>
<html>
<head>

<title>Obras Sociales</title>


<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<link href="../ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="../shared/js/fn_windows.js"></script>
<script src="../shared/js/fn_confirm.js"></script>
<script src="../shared/js/fn_ayuda.js"></script>

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_locales(){
	if (Trim(document.datos_02.descripcion.value) == ""){
		abrirAlert("Debe ingresar la Descripci&oacute;n de la Obra Social.");
		document.datos_02.descripcion.focus();
		return false;
		}
	else{
		return true;
	}
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'obrassocialesV2_06.asp'		//url_valid_06
						,'obrassocialesV2_03.asp'		//url_AM
						,'dialogAlert'							//id_dialogAlert
						,'datos_02'								//id_form_datos
						,window.parent.ifrm.location			//location_reload
						,Validaciones_locales					//funcion_Validaciones_locales
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete"				//id_dialogConfirmDelete
																	,"obrassocialesV2_04.asp"	//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"obrassocialesV2_06.asp"				//url_valid_06
														,"obrassocialesV2_03.asp"				//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
	    <tr>
			<td class="title">
				Obras Sociales
            </td>
			<td class="title"> </td>			
        </tr>
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra"></td>
		  
          <td nowrap align="right" class="barra">          		  
			<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','obrassocialesV2_02.asp?Tipo=A',600,300)"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta"></a>          
		  </td>
		  
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe  id="ifrm" name="ifrm" src="obrassocialesV2_01.asp" width="100%" height="100%"></iframe> 		  
	      </td>
        </tr>
		
      </table>
		
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->		
		<div id="dialog" title="Obras Sociales"> 			</div>	  
		
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>
</html>
