<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
'Archivo: cajamovimientos_con_00.asp
'Descripción: Administración de cajamovimientos
'Autor : Trivoli
'Fecha: 31/05/2015

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql

dim l_p_carga_js
Dim l_p_id_venta
Dim l_p_id_compra

l_p_carga_js = request.querystring("p_carga_js")
l_p_id_venta = request.querystring("p_id_venta")
l_p_id_compra = request.querystring("p_id_compra")
'response.write  "p_id_compra "&l_p_id_compra&"</br>"
  %>

<html>
<head>

<title>Administracion de Caja Movimientos</title>
<% if l_p_carga_js = "S" or isnull(l_p_carga_js) or l_p_carga_js = "" then%>
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ayuda.js"></script>
<script src="/trivoliSwimming/shared/js/fn_fechas.js"></script>


<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>
<% end if %>
<script src="/trivoliSwimming/shared/js/fn_numeros.js"></script>
<!-- Comienzo Datepicker -->
<script>
$(function () {

		
$( "#filt_fechadesde_cm" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar16.png",
	buttonImageOnly: true
});

$( "#filt_fechahasta_cm" ).datepicker({
	showOn: "button",
	buttonImage: "/trivoliSwimming/shared/images/calendar16.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->
<script>
function Validaciones_locales_mc(){

	if (document.datos_02_mc.fecha.value == ""){
		alert("Debe ingresar la fecha del movimiento.");
		document.datos_02_mc.fecha.focus();
		return false;
	}	
	if (document.datos_02_mc.tipoes.value == ""){
		alert("Debe ingresar el Tipo.");
		document.datos_02_mc.tipoes.focus();
		return false;
	}	
	
	if (document.datos_02_mc.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02_mc.idtipomovimiento.focus();
		return false;
	}	
	
	/*if (document.datos_02_mc.detalle.value == ""){
		alert("Debe ingresar el Detalle.");
		document.datos_02_mc.detalle.focus();
		return false;
	}*/		
	
	if (document.datos_02_mc.idunidadnegocio.value == "0"){
		alert("Debe ingresar la Unidad de Negocio.");
		document.datos_02_mc.idunidadnegocio.focus();
		return false;
	}	
	
	if (document.datos_02_mc.idmediopago.value == "0"){
		alert("Debe ingresar el Medio de Pago.");
		document.datos_02_mc.idmediopago.focus();
		return false;
	}	
	
	if (document.datos_02_mc.mediodepagocheque.value == document.datos_02_mc.idmediopago.value){
		if (document.datos_02_mc.idcheque.value == "0"){
			alert("Debe ingresar el Cheque.");
			document.datos_02_mc.idcheque.focus();
			return false;
		}	

	}		
	
	if (document.datos_02_mc.idtipomovimiento.value == "0"){
		alert("Debe ingresar el Tipo de Movimiento.");
		document.datos_02_mc.idtipomovimiento.focus();
		return false;
	}				

	if (document.datos_02_mc.monto.value == ""){
		alert("Debe ingresar un Monto.");
		document.datos_02_mc.monto.focus();
		return false;
	}		
	
	if (document.datos_02_mc.monto.value == "0"){
		alert("El Monto debe ser distinto de Cero.");
		document.datos_02_mc.monto.focus();
		return false;
	}		
	
	document.datos_02_mc.monto2.value = document.datos_02_mc.monto.value.replace(",", ".");
	if (!validanumero(document.datos_02_mc.monto2, 15, 4)){
		  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
		  document.datos.monto.focus();
		  document.datos.monto.select();
		  return;
	}		
	
	if (document.datos_02_mc.idresponsable.value == "0"){
		alert("Debe ingresar el Responsable.");
		document.datos_02_mc.idresponsable.focus();
		return false;
	}	


	return true;
}

function Submit_Formulario_mc() {
	Validar_Formulario(	'dialog_mc'								//id_dialog
						,'cajamovimientos_con_06.asp'					//url_valid_06
						,'cajamovimientos_con_03.asp'					//url_AM
						,'dialogAlert_mc'							//id_dialogAlert
						,'datos_02_mc'								//id_form_datos
						,null //window.parent.ifrm_mc.location			//location_reload
						,Validaciones_locales_mc					//funcion_Validaciones_locales
						,"ifrm_mc"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_mc"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete_mc"				//id_dialogConfirmDelete
																	,"cajamovimientos_con_04.asp"				//url_baja
																	,"dialogAlert_mc"						//id_dialogAlert
																	,"detalle_01_mc"						//id_form_datos
																	,"ifrm_mc"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog_mc" 										//id_dialog
														,"cajamovimientos_con_06.asp"							//url_valid_06
														,"cajamovimientos_con_03.asp"							//url_AM
														,"dialogAlert_mc"									//id_dialogAlert	
														,"datos_02_mc"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_mc							//funcion_Validaciones_locales	
														,"ifrm_mc"											//id_ifrm_form_datos														
														); 
								inicializar_dialogoContenedor(	"dialog_cont_BusqCompraOrigen" 										//id_dialog
																); 				
								inicializar_dialogoContenedor(	"dialog_cont_BusqVentaOrigen" 										//id_dialog
																); 																												
							});
							
</script>
<!--	FIN VENTANAS MODALES    -->

<script>

function Buscar_mc(){

	$("#filtro_00_mc").val("");

	// Nombre
	if ($("#filt_detalle_cm").val() != 0){
		$("#filtro_00_mc").val(" cajamovimientos.detalle like '*" + $("#filt_detalle_cm").val() + "*'");
	}		
    
	//Fecha desde
	if ($("#filt_fechadesde_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+ " cajamovimientos.fecha  >= " + cambiafechaYYYYMMDD($("#filt_fechadesde_cm").val(),true,1)
							);		
	}	

	//Fecha hasta
	if ($("#filt_fechahasta_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+ " cajamovimientos.fecha  <= " + cambiafechaYYYYMMDD($("#filt_fechahasta_cm").val(),true,1)
							);		
	}	
	
	//Numero de cheque 
	if ($("#filt_nrocheque_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+" cheques.numero like '*" + $("#filt_nrocheque_cm").val() + "*'"
							);		
	}	
	
	//banco 
	if ($("#filt_banco_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+" bancos.nombre_banco like '*" + $("#filt_banco_cm").val() + "*'"
							);		
	}
	
	//Cliente/Proveedor 
	if ($("#filt_cliprov_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+"( proveedores.nombre like '*" + $("#filt_cliprov_cm").val() + "*'"
								+" or clientes.nombre like '*" + $("#filt_cliprov_cm").val() + "*')"
							);		
	}
	
	
	//Medio de pago 
	if ($("#filt_idmediopago_cm").val() != 0){
		if ($("#filtro_00_mc").val() != 0){
			$("#filtro_00_mc").val( $("#filtro_00_mc").val() + " and ");
		}
		$("#filtro_00_mc").val(
								$("#filtro_00_mc").val() 
								+" cajamovimientos.idmediopago = " + $("#filt_idmediopago_cm").val() 
							);		
	}	
	window.ifrm_mc.location = 'cajamovimientos_con_01.asp?filtro=' + $("#filtro_00_mc").val()+'&p_id_venta='+"<%= l_p_id_venta%>"+'&p_id_compra='+"<%= l_p_id_compra%>";
}

function Limpiar_mc(){
	window.ifrm_mc.location = 'cajamovimientos_con_01.asp?p_id_venta='+"<%= l_p_id_venta%>"+'&p_id_compra='+"<%= l_p_id_compra%>";
}

function BuscarCompraOrigen(){	
	<%if l_p_id_venta <> "" or l_p_id_compra <> "" then%>
			alert("No se puede seleccionar otra operacion. Vaya a la pantalla de Caja");
	<%else%>	
		abrirDialogo('dialog_cont_BusqCompraOrigen','BuscarCompraOrigenV2_00.asp?Tipo=A&Alta=N&fn_asign_pac=volver_AsignarCompraOrigen&dnioblig=N&hcoblig=N',900,250);			
	<%end if%>	
	
}

function BuscarVentaOrigen(){	
	
	<%if l_p_id_venta <> "" or l_p_id_compra <> "" then%>
			alert("No se puede seleccionar otra operacion. Vaya a la pantalla de Caja");
	<%else%>	
		abrirDialogo('dialog_cont_BusqVentaOrigen','BuscarVentaOrigenV2_00.asp?Tipo=A&Alta=N&fn_asign_pac=volver_AsignarVentaOrigen&dnioblig=N&hcoblig=N',900,250);		
	<%end if%>		
}

function volver_AsignarCompraOrigen(id, fecha,  nombre){

	document.datos_02_mc.compraorigen.value = nombre + " - " + fecha ;
	document.datos_02_mc.idcompraorigen.value = id;
	
	$("#dialog_cont_BusqCompraOrigen").dialog("close");
}

function volver_AsignarVentaOrigen(id, fecha,  nombre){

	document.datos_02_mc.ventaorigen.value = nombre + " - " + fecha ;
	document.datos_02_mc.idventaorigen.value = id;
	
	$("#dialog_cont_BusqVentaOrigen").dialog("close");
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
                            Administracion de Caja
                        </td>
                    </tr>
                </table>
    		</td>
        </tr>
        <tr>
			<td>
                <input type="hidden" id="filtro_00_mc" name="filtro_00_mc" value="">
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
					    <td>
							<b>Cliente: </b><input  type="text" id="filt_cliprov_cm" name="filt_cliprov_cm" size="21" maxlength="21" value="" >
						</td>
					    <td>
							<b>Fecha: </b><input id="filt_fechadesde_cm" type="text" name="filt_fechadesde_cm" size="10" maxlength="10" value="" >							
						</td>
						<td>
							<b>Nro Cheque: </b><input id="filt_nrocheque_cm" type="text" name="filt_nrocheque_cm" size="10" maxlength="10" value="" >
						</td>

						<td>
							<b>Medio de pago: </b>
							<select name="filt_idmediopago_cm" id="filt_idmediopago_cm" size="1" style="width:100;" >
								<option value="0" selected>&nbsp;Todos</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM mediosdepago "
								l_sql = l_sql & " where mediosdepago.empnro = " & Session("empnro")   								
								l_sql  = l_sql  & " ORDER BY titulo "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value=<%= l_rs("id") %> > 
								<%= l_rs("titulo") %>  </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>								
							</select>	
						</td>						
						
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar_mc();" ><img  src="/trivoliSwimming/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar_mc();" ><img  src="/trivoliSwimming/shared/images/Limpiar_24.png" border="0" title="Limpiar">                            
							<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialog_mc','cajamovimientos_con_02.asp?Tipo=A&p_id_venta='+'<%= l_p_id_venta%>'+'&p_id_compra='+'<%= l_p_id_compra%>',650,450)"><img  src="/trivoliSwimming/shared/images/Agregar_24.png" border="0" title="Agregar Movimiento"></a>    
                        </td>
                    </tr>
					<tr>
						<td>
							<b>Detalle: </b><input  type="text" id="filt_detalle_cm" name="filt_detalle_cm" size="21" maxlength="21" value="" >
						</td>
					    <td>					
							<b>Hasta: </b><input id="filt_fechahasta_cm" type="text" name="filt_fechahasta_cm" size="10" maxlength="10" value="" >	
						</td>						
						<td>
							<b>Banco&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp: </b><input id="filt_banco_cm" type="text" name="filt_banco_cm" size="10" maxlength="10" value="" >	
						</td>
						<td></td>
						<td></td>
					</tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe id="ifrm_mc" name="ifrm_mc" src="" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
	
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialog_mc" title="Cajas"> 			</div>	  
				
		<div id="dialogAlert_mc" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete_mc" title="Consulta">		</div>		
		
		<div id="dialog_cont_BusqCompraOrigen" title="Buscar Compra Origen">		</div>		
		
		<div id="dialog_cont_BusqVentaOrigen" title="Buscar Venta Origen">		</div>				
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>

<script>
	Buscar_mc();
</script>
</html>
