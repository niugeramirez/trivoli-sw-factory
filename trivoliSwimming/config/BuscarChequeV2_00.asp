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
  
  
%>
<html>
<head>
<link href="/trivoliSwimming/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Buscar Ventas</title>
<!--	VENTANAS MODALES        -->
<script>

function AsignarCheque(id, fecha, numero, banco, importe){

	<%= l_fn_asign_pac %>(id, fecha, numero, banco, importe);
}

function devolver_cheque_editado(id){

	<%= l_fn_asign_pac %>(	id,
							document.datos_02_cheq.fecha_emision.value, 
							document.datos_02_cheq.numero.value, 
							$( "#idbanco option:selected" ).text().trim(),
							document.datos_02_cheq.importe.value  
							);
}

function Validaciones_locales_EditCheq(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_cheq();
}

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_BusqEdicCheq"									//id_dialogAlert
														);

								inicializar_dialogoABM(	"dialog_cont_EditCheq" 										//id_dialog
														,"cheques_con_06.asp"				//url_valid_06
														,"cheques_con_03.asp"				//url_AM
														,"dialogAlert_BusqEdicCheq"									//id_dialogAlert															
														,"datos_02_cheq"										//id_form_datos															
														,null //window.parent.ifrm.location					//location_reload														
														,Validaciones_locales_EditCheq							//funcion_Validaciones_locales	
														,"ifrm_mc"											//id_ifrm_form_datos	
														,devolver_cheque_editado //fn_post_AM														
														); 																
							});
</script>
<!--	FIN VENTANAS MODALES    -->
<script>

function Buscar(){
	
	$("#filtro").val("");

	// Nombre
	if ($("#nombre").val() != 0){
		//Eugenio 02/05/2016 en lugar de enviar * para los comodines envio **, con esto el el 01 reemplazo el ** por %
		//Esto es porque el otro filtro implica una operacion matematica de multiplicacion, con lo que se modifica el * y hace un calculo erroneo	
		$("#filtro").val(" cheques.numero like '**" + $("#nombre").val() + "**'");
	}					
				
	//Estado 
	if ($("#filt_estado_cheq_busq").val() != 0){
		if ($("#filtro").val() != 0){
			$("#filtro").val( $("#filtro").val() + " and ");
		}
		$("#filtro").val(
								$("#filtro").val() 
								+" dbo.get_estado_cheque(cheques.id) = " + $("#filt_estado_cheq_busq").val() 
							);		
	} 
	
	window.ifrm_BusqChe.location = 'BuscarChequeV2_01.asp?asistente=0&filtro=' + document.form_BusqChe.filtro.value;

}


function AltaCheque(){

	abrirDialogo('dialog_cont_EditCheq','cheques_con_02.asp?Tipo=A&ventana=3&dni=<%= l_dni %>&hcoblig=<%= l_hcoblig %>',600,300);
}

function Limpiar(){
	window.ifrm_BusqChe.location = 'BuscarChequeV2_01.asp?asistente=0&filtro=1=2' ;
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.form_BusqCli.nombre.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">	
		<tr>
			<td align="center" colspan="2">
				
					<form name="form_BusqChe" id="form_BusqChe" action="nada" target="nada">
						<input type="hidden" name="filtro" id="filtro" value="">
						<table border="0">
							<tr>
								<td align="right"><b>Numero: </b></td>
								<td align="left" colspan="3" ><input  type="text" name="nombre" id="nombre" size="21" maxlength="21" value="" ></td>					
								<td> <b>Estado: </b>
									<select name="filt_estado_cheq_busq" id="filt_estado_cheq_busq" size="1" style="width:100;" >
										<option value="0" selected>&nbsp;Todos</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM cheques_estado "
										rsOpen l_rs, cn, l_sql, 0
										response.write l_sql
										do until l_rs.eof		%>	
										<option value=<%= l_rs("id") %> > 
										<%= l_rs("estado") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>								
									</select>							
								</td>
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
					<a href="Javascript:AltaCheque();"><img src="../shared/images/Agregar_24.png" border="0" title="Alta Cliente"></a>	
				<% End If %>
			</td>			
		</tr>				
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
			<iframe scrolling="yes" name="ifrm_BusqChe" id="ifrm_BusqCli"  src="BuscarChequeV2_01.asp?asistente=0&filtro=1=1" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		

      </table>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->		
		<div id="dialogAlert_BusqEdicCheq" title="Mensaje">				</div>			
		<div id="dialog_cont_EditCli" title="Editar Clientes">		</div>	
		<div id="dialog_cont_EditCheq" title="Editar Cheque">		</div>	
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>
</html>
