<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/ess/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Descripción:;"
  l_Campos    = "agedes"
  l_Tipos     = "T;"

' Orden
  l_Orden     = "Descripción:;"
  l_CamposOr  = "agedes"

  Dim l_id
  
  l_id = request.querystring("id")
  Dim l_rs
  Dim l_sql
  
  Dim l_obra_social
  
  
  
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT * "
  l_sql = l_sql & " FROM obrassociales "
  'l_sql = l_sql & " INNER JOIN obrassociales ON obrassociales.id = listaprecioscabecera.idobrasocial "
  l_sql = l_sql & " WHERE obrassociales.id = " & l_id
  rsOpen l_rs, cn, l_sql, 0 
  if not l_rs.eof then
  	l_obra_social = l_rs("descripcion")
  else
  	l_obra_social = ""
  	
  end if
  
  
  
%>
<html>
<head>

<title>Lista de Precios</title>
<link href="../ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<!--	VENTANAS MODALES        -->
<!-- <script src="../js/ventanas_modales_custom_V2.js"></script> -->

<script>
function Validaciones_locales_lista(){
	if ((document.datos_02_LP.fecha.value == "")&&(!validarfecha(document.datos_02_LP.fecha))){
		 document.datos_02_LP.fecha.focus();
		 return false;
	}


	if (Trim(document.datos_02_LP.titulo.value) == ""){
		alert("Debe ingresar el Titulo.");
		document.datos_02_LP.titulo.focus();
		return false;
	}

	return true;
}

function Submit_Formulario_lista() {
	Validar_Formulario(	'dialog_lista'								//id_dialog
						,'listadeprecios_con_06.asp'		//url_valid_06
						,'listadeprecios_con_03.asp'		//url_AM
						,'dialogAlert_lista'							//id_dialogAlert
						,'datos_02_LP'								//id_form_datos
						,null //window.parent.ifrm_01_LP.location			//location_reload
						,Validaciones_locales_lista					//funcion_Validaciones_locales
						,"ifrm_01_LP"											//id_ifrm_form_datos
					);
} 

$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlert_lista"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDelete_lista"				//id_dialogConfirmDelete
																	,"listadeprecios_con_04.asp"	//url_baja
																	,"dialogAlert_lista"						//id_dialogAlert
																	,"detalle_01_LP"						//id_form_datos
																	,"ifrm_01_LP"								//id_ifrm_form_datos
																	,null //window.parent.ifrm_01_LP.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialog_lista" 										//id_dialog
														,"listadeprecios_con_06.asp"				//url_valid_06
														,"listadeprecios_con_03.asp"				//url_AM
														,"dialogAlert_lista"									//id_dialogAlert	
														,"datos_02_LP"										//id_form_datos		
														,null //window.parent.ifrm_01_LP.location					//location_reload
														,Validaciones_locales_lista						//funcion_Validaciones_locales	
														,"ifrm_01_LP"											//id_ifrm_form_datos														
														); 
													
							});
</script>
<!--	FIN VENTANAS MODALES    -->


</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">Lista de Precios</td>
          <td nowrap align="right" class="barra">          
			<a class="sidebtnABM" href="Javascript:abrirDialogo('dialog_lista','listadeprecios_con_02.asp?Tipo=A&idobrasocial=<%= l_id%>',520,350);" ><img  src="../shared/images/Agregar_24.png" border="0" title="Alta">
		  </td>
        </tr>
		<tr>
		<td align="right"><b>Obra Social: </b></td>
			<td><input  type="text" name="legape" size="41" maxlength="21" value="<%= l_obra_social %>" readonly   class="deshabinp" >
			</td>
		</tr>				
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="Yes" id="ifrm_01_LP" name="ifrm_01_LP" src="listadeprecios_con_01.asp?idobrasocial=<%= l_id  %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
      </table>

		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->		
		<div id="dialog_lista" title="Lista de Precios"> 			</div>	  
		
		<div id="dialogAlert_lista" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDelete_lista" title="Consulta">		</div>		
		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->		  
</body>
</html>
