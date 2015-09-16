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


<script>
//Esto va antes de la importacion de ventanas_modales_custom.js
function Validaciones_locales(){
	if (Trim(document.datos.descripcion.value) == ""){
		abrirAlert("Debe ingresar la Descripci&oacute;n de la Obra Social.");
		document.datos.descripcion.focus();
		return false;
		}
	else{
		return true;
	}
}
</script>
<script src="../js/ventanas_modales_custom.js"></script>

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
			<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('obrassocialesV2_02.asp?Tipo=A')"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta"></a>          
		  </td>
		  
        </tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe  id="ifrm" name="ifrm" src="obrassocialesV2_01.asp" width="100%" height="100%"></iframe> 		  
	      </td>
        </tr>
		
      </table>
		
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->
		<!--	URL´s        -->
		<input type="hidden" id="url_AM" value="obrassocialesV2_03.asp">
		<input type="hidden" id="url_valid_06" value="obrassocialesV2_06.asp">	
		<input type="hidden" id="url_baja" value="0">	
		<input type="hidden" id="url_base_baja" value="obrassocialesV2_04.asp?">	
		
		<!--	DIV´s Dialogos       -->
		<input type="hidden" id="id_dialog" value="dialog">
		<input type="hidden" id="width_dialog" value="600">
		<input type="hidden" id="height_dialog" value="300">		
		<div id="dialog" title="Obras Sociales"> 			</div>	  
		
		<input type="hidden" id="id_dialogAlert" value="dialogAlert">
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<input type="hidden" id="id_dialogConfirmDelete" value="dialogConfirmDelete">
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		
</body>
</html>
