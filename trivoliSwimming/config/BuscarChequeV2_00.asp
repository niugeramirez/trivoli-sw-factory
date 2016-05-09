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
	<%= l_fn_asign_pac %>(id, fecha, numero, banco, importe );
}

function Validaciones_locales_EditPac(){
	//como la pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local en el 02, y se invoca desde la ventana llamadora
	return Validaciones_locales_EditPac_02()
}


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
				
	
	window.ifrm_BusqChe.location = 'BuscarChequeV2_01.asp?asistente=0&filtro=' + document.form_BusqChe.filtro.value;

}


function AltaCliente(){

	abrirDialogo('dialogHCR_cont_EditCli','EditarclientesV2_02.asp?Tipo=A&ventana=3&dni=<%= l_dni %>&hcoblig=<%= l_hcoblig %>',600,300);
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
		<div id="dialogAlert_BusqEdicCli" title="Mensaje">				</div>			
		<div id="dialog_cont_EditCli" title="Editar Clientes">		</div>	
		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>
</html>
