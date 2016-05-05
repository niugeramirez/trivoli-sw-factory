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

<script src="js_pantallas/cheques.js"></script>

<!-- Comienzo Datepicker -->
<script>
$(function () {

		
$( "#filt_fechavcntodesde_cheq" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar16.png",
	buttonImageOnly: true
});

$( "#filt_fechavcntohasta_cheq" ).datepicker({
	showOn: "button",
	buttonImage: "/trivoliSwimming/shared/images/calendar16.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>


$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_cheq"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete_cheq"				//id_dialogConfirmDelete
																	,"cheques_con_04.asp"				//url_baja
																	,"dialogAlert_cheq"						//id_dialogAlert
																	,"detalle_01_cheq"						//id_form_datos
																	,"ifrm_cheq"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog_cheq" 										//id_dialog
														,"cheques_con_06.asp"							//url_valid_06
														,"cheques_con_03.asp"							//url_AM
														,"dialogAlert_cheq"									//id_dialogAlert	
														,"datos_02_cheq"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_cheq							//funcion_Validaciones_locales	
														,"ifrm_cheq"											//id_ifrm_form_datos														
														); 
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar_cheq(){

	$("#filtro_00_cheq").val("");

	// nro cheque
	if ($("#inpnombre_cheq").val() != 0){
		$("#filtro_00_cheq").val(" cheques.numero like '*" + $("#inpnombre_cheq").val() + "*'");
	}		
	//Fecha vencimiento desde
	if ($("#filt_fechavcntodesde_cheq").val() != 0){
		if ($("#filtro_00_cheq").val() != 0){
			$("#filtro_00_cheq").val( $("#filtro_00_cheq").val() + " and ");
		}
		$("#filtro_00_cheq").val(
								$("#filtro_00_cheq").val() 
								+ " cheques.fecha_vencimiento  >= " + cambiafechaYYYYMMDD($("#filt_fechavcntodesde_cheq").val(),true,1)
							);		
	}	

	//Fechavencimiento hasta
	if ($("#filt_fechavcntohasta_cheq").val() != 0){
		if ($("#filtro_00_cheq").val() != 0){
			$("#filtro_00_cheq").val( $("#filtro_00_cheq").val() + " and ");
		}
		$("#filtro_00_cheq").val(
								$("#filtro_00_cheq").val() 
								+ " cheques.fecha_vencimiento  <= " + cambiafechaYYYYMMDD($("#filt_fechavcntohasta_cheq").val(),true,1)
							);		
	}	    
	window.ifrm_cheq.location = 'cheques_con_01.asp?asistente=0&filtro=' + $("#filtro_00_cheq").val();
}

function Limpiar_cheq(){
	window.ifrm_cheq.location = 'cheques_con_01.asp';
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
                <input type="hidden" id="filtro_00_cheq" name="filtro_00_cheq" value="">
				<table border="0" width="100%">
                    <colgroup>
                        <col class="colWidth20">
                        <col class="colWidth20">
                        <col class="colWidth20">
                        <col class="colWidth20">
						<col class="colWidth20">
                    </colgroup>
                    <tbody>
				    <tr>
					    <td><b>Numero: </b><input  type="text" id="inpnombre_cheq" name="inpnombre_cheq" size="21" maxlength="21" value="" ></td>
						<td>
							<b>Fecha Vcnto: </b><input id="filt_fechavcntodesde_cheq" type="text" name="filt_fechavcntodesde_cheq" size="10" maxlength="10" value="" >							
						</td>
					    <td></td>
						<td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar_cheq();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar_cheq();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog_cheq','cheques_con_02.asp?Tipo=A',650,350)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Cliente"></a>    
                        </td>
                    </tr>
					<tr>
						<td></td>
					    <td>					
							<b>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbspHasta: </b><input id="filt_fechavcntohasta_cheq" type="text" name="filt_fechavcntodesde_cheq" size="10" maxlength="10" value="" >
						</td>						
						<td></td>
						<td></td>
						<td></td>
					</tr>					
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm_cheq" name="ifrm_cheq" src="clientes_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog_cheq" title="Cheques"> 			</div>	  
				
		<div id="dialogAlert_cheq" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete_cheq" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar_cheq();
</script>
</html>
