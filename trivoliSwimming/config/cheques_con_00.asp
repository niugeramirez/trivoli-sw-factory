<% Option Explicit %>
<% 
'Archivo: cheques_con_00.asp
'Descripción: Administración de cheques
'Autor : Trivoli
'Fecha: 31/05/2015

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql
  %>

<html>
<head>

<title>Administracion de Cheques</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>
<script src="/trivoliSwimming/shared/js/fn_numeros.js"></script>

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>
function Validaciones_locales(){

	
	if (document.datos_02.numero.value == ""){
		alert("Debe ingresar un Numero.");
		document.datos_02.numero.focus();
		return false;
	}	

	if (document.datos_02.idbanco.value == "0"){
		alert("Debe ingresar un Banco.");
		document.datos_02.idbanco.focus();
		return false;
	}

	if (document.datos_02.importe.value == ""){
		alert("Debe ingresar un Importe.");
		document.datos_02.importe.focus();
		return false;
	}		
	
	if (document.datos_02.importe.value == "0"){
		alert("El Importe debe ser distinto de Cero.");
		document.datos_02.importe.focus();
		return false;
	}		
	
	document.datos_02.importe2.value = document.datos_02.importe.value.replace(",", ".");
	if (!validanumero(document.datos_02.importe2, 15, 4)){
		  alert("El Importe no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.importe.focus();
		  document.datos.importe.select();
		  return;
	}		
	
	if (document.datos_02.emisor.value == "0"){
		alert("Debe ingresar un Emisor.");
		document.datos_02.emisor.focus();
		return false;
	}	

/*
	if (document.datos_02.nombre.value == ""){
		alert("Debe ingresar el Nombre del Cliente.");
		document.datos_02.nombre.focus();
		return false;
	}

	if (document.datos_02.idtemplatereserva.value == 0){
		alert("Debe ingresar el Modelo.");
		document.datos_02.idtemplatereserva.focus();
		return false;
	}

	if (document.datos_02.cantturnossimult.value == ""){
		alert("Debe ingresar la Cantidad de Turnos Simultaneos.");
		document.datos_02.cantturnossimult.focus();
		return false;
	}
	*/
	return true;
}

function Submit_Formulario() {
	Validar_Formulario(	'dialog'								//id_dialog
						,'cheques_con_06.asp'					//url_valid_06
						,'cheques_con_03.asp'					//url_AM
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
																	,"cheques_con_04.asp"				//url_baja
																	,"dialogAlert"						//id_dialogAlert
																	,"detalle_01"						//id_form_datos
																	,"ifrm"								//id_ifrm_form_datos
																	,window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog" 										//id_dialog
														,"cheques_con_06.asp"							//url_valid_06
														,"cheques_con_03.asp"							//url_AM
														,"dialogAlert"									//id_dialogAlert	
														,"datos_02"										//id_form_datos		
														,window.parent.ifrm.location					//location_reload
														,Validaciones_locales							//funcion_Validaciones_locales														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar(){

	$("#filtro_00").val("");

	// Nombre
	if ($("#inpnombre").val() != 0){
		$("#filtro_00").val(" cheques.numero like '*" + $("#inpnombre").val() + "*'");
	}		
    
	window.ifrm.location = 'cheques_con_01.asp?asistente=0&filtro=' + $("#filtro_00").val();
}

function Limpiar(){
	window.ifrm.location = 'cheques_con_01.asp';
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
                            Administracion de Cheques
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
					    <td><b>Numero: </b></td>
						<td><input  type="text" id="inpnombre" name="inpnombre" size="21" maxlength="21" value="" ></td>
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog','cheques_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Cliente"></a>    
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm" name="ifrm" src="clientes_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog" title="Cheques"> 			</div>	  
				
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar();
</script>
</html>
