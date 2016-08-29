<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_00.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 27/11/2007

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Descripción:;Area"
  l_Campos    = "coudes;aredes"
  l_Tipos     = "T;T"

' Orden
  l_Orden     = "P/S:;Ctr Num:;Date:;Client:;Quality:;Volumen:;Port:;Term:;Company:;Product:"
  l_CamposOr  = "conpursal;ctrnum;confec;clidesabr;quadesabr;conquantity;pordes;terdes;comdesabr;prodesabr"

  Dim l_rs
  Dim l_sql
  
  Dim l_dia
  Dim l_mes
  Dim l_anio
  Dim l_id
  Dim l_fecha
  
  Dim l_buscar
  
l_dia = Request.Querystring("day")  
l_mes = Request.Querystring("Month")
l_anio = Request.Querystring("Year")
l_id = Request.Querystring("id")

If IsEmpty(l_dia) then 
	l_fecha = date()
else
	l_fecha = cstr(l_dia) & "/" & cstr(l_mes) & "/" & cstr(l_anio)
end if 

l_buscar = true
If IsEmpty(l_id) then   
	l_buscar = false
	l_id = 0
end if
%>
<html>
<head>
<title>Visitas</title>
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<script src="../js/jquery.ui.datepicker-es.js"></script>

<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<script src="js_pantallas/visitas.js"></script>

<!--	VENTANAS MODALES        -->
<script src="../js/ventanas_modales_custom_V2.js"></script>

<script>




$(document).ready(function() { 
								inicializar_dialogAlert("dialogAlertVisitas"									//id_dialogAlert
														);
								inicializar_dialogConfirmDelete(	"dialogConfirmDeleteVisitas"				//id_dialogConfirmDelete
																	,"VisitasV2_con_04.asp"				//url_baja
																	,"dialogAlertVisitas"						//id_dialogAlert
																	,"detalle_01_visit"						//id_form_datos
																	,"ifrm_visit"								//id_ifrm_form_datos
																	,null //window.parent.ifrm.location		//location_reload
																	);
								inicializar_dialogoABM(	"dialogVisit" 										//id_dialog
														,"visitasV2_con_06.asp"							//url_valid_06
														,"visitasV2_con_03.asp"							//url_AM
														,"dialogAlertVisitas"									//id_dialogAlert	
														,"datosAVST"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_AVST							//funcion_Validaciones_locales	
														,"ifrm_visit"											//id_ifrm_form_datos														
														); 

								inicializar_dialogoABM(	"dialogVisitCT" 										//id_dialog
														,"visitasV2_conturno_con_06.asp"							//url_valid_06
														,"visitasV2_conturno_con_03.asp"							//url_AM
														,"dialogAlertVisitas"									//id_dialogAlert	
														,"datos_02_AVCT"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_AVconT							//funcion_Validaciones_locales	
														,"ifrm_visit"											//id_ifrm_form_datos														
														); 														
								inicializar_dialogoContenedor(	"dialog_cont_Pagos" 										//id_dialog
																); 
								//esta linea la agrego solo para refrescar cuando se cierra el dialogo contenedor, se podría parametrizar de modo de recibir
								//la funcion como parametro que se debe ejecutar al
								$( "#dialog_cont_Pagos" ).dialog({
									close: function () {$(this).empty(); Buscar_Visitas();}
								});			
																
							});
</script>
<!--	FIN VENTANAS MODALES    -->

<!-- Comienzo Datepicker -->
<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#fechadesde" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});


});
</script>
<!-- Final Datepicker -->

<script>




function Buscar_Visitas(){
	var tieneotro;
	var estado;
	document.datos.filtro_visit.value = "";
	tieneotro = "no";
	estado = "si";

	// fec. desde
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro_visit.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			document.datos.filtro_visit.value += " CONVERT(VARCHAR(10), visitas.fecha, 101)  = " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
			tieneotro = "si";
		}else{
			estado = "no";
		}
	}

	
	if (document.datos.id.value == "0"){
		alert("Debe ingresar un Medico.");
		document.datos.id.focus();
		return;
	}	
	
	
	if (document.datos.id.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro_visit.value += " AND visitas.idrecursoreservable = " + document.datos.id.value + "";
		}else{
			document.datos.filtro_visit.value += " visitas.idrecursoreservable = " + document.datos.id.value + "";
		}
		tieneotro = "si";
	}	

	if (estado == "si"){
		window.ifrm_visit.location = 'visitasV2_con_01.asp?idrecursoreservable=' + document.datos.id.value + '&filtro=' + document.datos.filtro_visit.value;
	}
}



function Limpiar(){

	document.datos.fechadesde.value = "<%= date() - 1 %>";
	//document.datos.fechahasta.value = "<%= date() %>";

	document.datos.id.value     = 0;

	window.ifrm_visit.location = 'visitas_con_01.asp';
}




function AltaVisita(){ 

	if (document.datos.id.value == 0) {
		alert("Debe seleccionar un Medico")
		return;
	}		
	else {		
		abrirDialogo('dialogVisit','visitasV2_con_02.asp?Tipo=A&fechadesde='+ document.datos.fechadesde.value + '&idrecursoreservable='+ document.datos.id.value,950,350);
	}		
	
}

function AltaVisitaconTurno(){ 

	if (document.datos.id.value == 0) {
		alert("Debe seleccionar un Medico")
		return;
	}		
	else {
		//abrirVentana("altavisitaconturno_con_00.asp?Tipo=A&fechadesde="+ document.datos.fechadesde.value + "&idrecursoreservable="+ document.datos.id.value,'',800,600);
		abrirDialogo('dialogVisitCT',"visitasV2_conturno_con_02.asp?Tipo=A&fechadesde="+ document.datos.fechadesde.value + "&idrecursoreservable="+ document.datos.id.value,800,450);
	}		
	
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fechadesde.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">		
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<form name="datos">
					<input type="hidden" name="filtro_visit" value="">

					<tr>
						<td align="right"><b>Fecha: </b></td>
						<td><input  type="text" id="fechadesde"  name="fechadesde" size="10" maxlength="10" value="<%= l_fecha%>" ></td>
						<td  align="right" nowrap><b>M&eacute;dico: </b></td>
						<td><select name="id" size="1" style="width:200;">
								<option value=0 selected>Seleccionar un M&eacute;dico</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables  "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.id.value= "<%= l_id %>"</script>
						</td>			
						<td></td>
						<td align="center">
							<a class="sidebtnABM" href="Javascript:Buscar_Visitas();" ><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar">
							<a class="sidebtnABM" href="Javascript:AltaVisita();" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agregar Visitas sin Turno">
							<a class="sidebtnABM" href="Javascript:AltaVisitaconTurno();"><img  src="/turnos/shared/images/Add-Appointment-icon_24.png" border="0" title="Agregar Visitas con Turno">
						</td>					
								
					</tr>



					<tr>

						
					</tr>					

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm_visit" id="ifrm_visit" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
			</form>		
      </table>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->				
		<div id="dialogVisit" title="Visitas"> 			</div>	  
		
		<div id="dialogVisitCT" title="Visitas con Turno"> 			</div>	
				
		<div id="dialogAlertVisitas" title="Mensaje">				</div>	
		
		<div id="dialogConfirmDeleteVisitas" title="Consulta">		</div>	
		<div id="dialog_cont_Pagos" title="Pagos">		</div>			
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->			  
</body>


</html>
