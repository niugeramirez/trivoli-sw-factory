<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: companies_con_02.asp
'Descripción: ABM de Companies
'Autor : Raul Chinestra
'Fecha: 26/11/2007

'Datos del formulario
Dim l_id
Dim l_titulo
Dim l_descripcion

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs


Dim l_horainicial 
Dim l_horafinal
Dim l_intervaloturnominutos
Dim l_calfec

Dim l_idrecursoreservable
Dim l_mediodepagoos
Dim l_osparticular




l_tipo = request.querystring("tipo")
l_idrecursoreservable = request.querystring("idrecursoreservable")
l_calfec  = request.querystring("fechadesde")


%>
<html>
<head>
<title>Agregar Visitas sin Turno</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
<!--
<script src="../js/jquery-1.8.0.js"></script>
<script src="../js/jquery-ui.js"></script>  
-->
<script src="../js/jquery.ui.datepicker-es.js"></script>
<link href="../ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<!--Inicio Autocomplete derivador -->

  <style>
  .ui-autocomplete-loading {
    background: white url("../images/ui-anim_basic_16x16.gif") right center no-repeat;
  }
.ui-autocomplete {
    max-height: 100px;
    overflow-y: auto;
    /* prevent horizontal scrollbar */
    overflow-x: hidden;
  }
  </style>
<script type="text/javascript" language="javascript">

    $(function() {
        $( "#med_derivador" ).bind( "keydown", function( event ) {
									if ( event.keyCode === $.ui.keyCode.TAB &&
										$( this ).autocomplete( "instance" ).menu.active ) {
									  event.preventDefault();
									}
								  })							  
		.autocomplete({
			source: function( request, response ) {
						  $.getJSON( "JSON_medicos_derivadores.asp", {
							term: request.term
						  }, response );
						},
            minLength: 0,
			select: function( event, ui ) {
					$( "#med_derivador" ).val( ui.item.label );
					$( "#idmedicoderivador" ).val( ui.item.id );			 
					return false;
				  },
			change: function(event,ui){
				  $(this).val((ui.item ? ui.item.label : ""));
				  ui.item ? $( "#idmedicoderivador" ).val( ui.item.id ) : $( "#idmedicoderivador" ).val( "" );
				}			  
		});
		
		$("#but_display_med_deirv")
			.button({
				  icons: {
					primary: "ui-icon-triangle-1-s"
				  },
				  text: false
			})
			.removeClass( "ui-corner-all" )
			.addClass( "ui-corner-right ui-button-icon" )
			.click(function() {
				$( "#med_derivador" ).autocomplete("search", "");
				$( "#med_derivador" ).focus();
			});
	
    });


</script>
<!--Fin Autocomplete derivador -->
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<!--Inicio ventanas modales-->
<script>
//Esto va antes de la importacion de ventanas_modales_custom.js
function Validaciones_locales(){
	//devuelvo siempre verdadero en este caso de modo de hacer los controles en el 06
	return true;
}
</script>
<script src="../js/ventanas_modales_custom.js"></script>
<!--Fin ventanas modales-->

<!-- Comienzo Datepicker -->

<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#calfec" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});



});
</script>
<!-- Final Datepicker -->

</head>
<script>
function Validar_Formulario_visita(){

	if (document.datos.pacienteid.value == "0"){
		alert("Debe ingresar el Paciente.");
		document.datos.pacienteid.focus();
		return;
	}

	if (document.datos.practicaid.value == "0"){
		alert("Debe ingresar la Practica.");
		document.datos.practicaid.focus();
		return;
	}

	document.datos.precio2.value = document.datos.precio.value.replace(",", ".");
	if (!validanumero(document.datos.precio2, 15, 4)){
			  alert("El Precio no es válido. Se permite hasta 15 enteros y 4 decimales.");	
			  document.datos.precio.focus();
			  document.datos.precio.select();
			  return;
	}

	if (document.datos.mediodepagoos.value == document.datos.idmediodepago.value)  {
		if (Trim(document.datos.idobrasocial.value) == "0"){
			alert("Debe ingresar la Obra Social.");
			document.datos.idobrasocial.focus();
			return;
		}
	}

	if (document.datos.importe.value == ""){
		if	(document.datos.idmediodepago.value != "0") {
			alert("Debe ingresar un Importe mayor o igual a 0.");
			document.datos.importe.focus();
			return;
		}
	}

	document.datos.importe2.value = document.datos.importe.value.replace(",", ".");

	if (!validanumero(document.datos.importe2, 15, 4)){
			  alert("El Monto no es válido. Se permite hasta 15 enteros y 4 decimales.");	
			  document.datos.importe.focus();
			  document.datos.importe.select();
			  return;
	}	

	if (document.datos.importe.value != 0)  {
		if (Trim(document.datos.idmediodepago.value) == "0"){
			alert("Debe ingresar el Medio de Pago.");
			document.datos.idmediodepago.focus();
			return;
		}
	}

	if (document.datos.idmediodepago.value != "0")  {
		if (Trim(document.datos.importe.value) == "0"){
			alert("Debe ingresar el Importe.");
			document.datos.importe.focus();
			return;
		}
	}


	var d=document.datos;
	document.valida.location = "altavisita_con_06.asp?pacienteid="+document.datos.pacienteid.value ; 


	valido_visita();

}

function valido_visita(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.agedes.focus();
}

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function EncontrePaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel, osid, os){
	document.datos.pacienteid.value = id;
	document.datos.apellido.value = apellido;
	document.datos.nombre.value = nombre;
	document.datos.nrohistoriaclinica.value = nrohistoriaclinica;
	document.datos.dni.value = dni;
	document.datos.domicilio.value = domicilio;
	document.datos.tel.value = tel;
	document.datos.osid.value = osid;
	document.datos.os.value = os;
	//document.datos.coudes.focus();
	document.datos.idobrasocial.value = osid;
	
	// lo dejo dividido asi por si mas adelante deshabilitamos algun control
	if (osid == document.datos.osparticular.value){
		document.datos.idobrasocial.value = 0;
		document.datos.idmediodepago.value = 0;
	}
    else {
		document.datos.idobrasocial.value = osid;
		document.datos.idmediodepago.value = document.datos.mediodepagoos.value;
	};
}

function EncontrePacienteAlta(id,apellido, nombre, dni,tel,domicilio,osid, os){
	
	document.datos.pacienteid.value = id;
	document.datos.apellido.value = apellido;
	document.datos.nombre.value = nombre;
	document.datos.dni.value = dni;
	document.datos.domicilio.value = domicilio;
	document.datos.tel.value = tel;
	document.datos.osid.value = osid;
	document.datos.os.value = os;
	
}

function BuscarPaciente(){
	abrirVentana('Buscarpacientes_con_00.asp?Tipo=A&Alta=S&dni=S&hc=S','',600,250);
}


function calcularprecio(){
	
	document.valida.location = "agregarpractica_con_06.asp?idos=" + document.datos.osid.value + "&practicaid="+ document.datos.practicaid.value ;	
}

function actualizarprecio(p_precio){	
	document.datos.precio.value = p_precio;
	
	// Si el medio de Pago es Obra social, copio el precio al importe
	if (document.datos.idmediodepago.value == document.datos.mediodepagoos.value ) { 
		document.datos.importe.value = p_precio;
	} 
 	else document.datos.importe.value = 0;	

}	


function ctrolmetodopago(){
	if (document.datos.mediodepagoos.value == document.datos.idmediodepago.value) {		
			document.datos.idobrasocial.disabled = false;							
		}
		else {
			document.datos.idobrasocial.disabled = true;							
			document.datos.idobrasocial.value = 0;	
		}	

}


</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'obtengo el Medio de Pago Obra Social
l_sql = "SELECT * "
l_sql = l_sql & " FROM mediosdepago "
l_sql  = l_sql  & " WHERE flag_obrasocial = -1 " 
l_sql = l_sql & " AND empnro = " & Session("empnro")
rsOpen l_rs, cn, l_sql, 0 
l_mediodepagoos = 0
if not l_rs.eof then
	l_mediodepagoos = l_rs("id")	
end if
l_rs.Close

l_sql = "SELECT  * "
l_sql  = l_sql  & " FROM obrassociales "
l_sql  = l_sql  & " WHERE isnull(obrassociales.flag_particular,0) = -1 "	
l_sql = l_sql & " AND empnro = " & Session("empnro")								
rsOpen l_rs, cn, l_sql, 0 
l_osparticular = 0
if not l_rs.eof then
	l_osparticular = l_rs("id")	
end if
l_rs.Close


select Case l_tipo
	Case "A":
		l_titulo = ""
		l_descripcion = ""
	Case "M":

end select
%>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" <% if l_tipo <> "M" then %> onload="javascript:BuscarPaciente();" <% End If %>>
<form name="datos" action="altavisita_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="idrecursoreservable" value="<%= l_idrecursoreservable %>">

<input type="hidden" name="id" value="<%= l_id %>">
<input type="hidden" name="pacienteid" value="0">

<input type="hidden" name="osid" value="<%= l_idobrasocial %>">

<input type="hidden" name="calfec" value="<%= l_calfec %>">


<input type="hidden" name="mediodepagoos" value="<%= l_mediodepagoos %>">
<input type="hidden" name="osparticular" value="<%= l_osparticular %>">



<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
	
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">						
					<tr>	
					<td colspan="4" align="center">
					<% if l_tipo <> "M" then %>
					<a href="Javascript:BuscarPaciente();"><img src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar Paciente"></a>		
					<% End If %>						

					</td>
					</tr>	
					<tr>
					    <td align="right"><b>Apellido:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="apellido" size="20" maxlength="20" value="<%= l_apellido %>">							
						</td>
					    <td align="right"><b>Nombre:</b></td>						
						<td>
							<input class="deshabinp" readonly="" type="text" name="nombre" size="20" maxlength="20" value="<%= l_nombre %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I.:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="dni" size="20" maxlength="20" value="<%= l_dni %>">
						</td>
					    <td align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
						</td>						
					</tr>
					<tr>
					    <td align="right"><b>Tel&eacute;fono:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="tel" size="20" maxlength="20" value="<%= l_tel %>">
						</td>
					    <td align="right"><b>Domicilio:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="domicilio" size="20" maxlength="20" value="<%= l_domicilio %>">
						</td>						
					</tr>
				
					<tr>
					    <td align="right"><b>Obra Social:</b></td>
						<td>
							<input class="deshabinp" readonly="" type="text" name="os" size="20" maxlength="20" value="<%= l_descripcion %>">
						</td>
						<td colspan="2" align="left">&nbsp;</td>
					    					
					</tr>	
					<tr>
						<td>&nbsp;
						</td>
					</tr>
						
					<tr>
						<td  align="right" nowrap><b>Practica (*): </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;" onchange="calcularprecio();">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql = l_sql & " WHERE empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.practicaid.value="0"</script>
						</td>					
					</tr>	
					
					<tr>
						<td  align="right" nowrap><b>Solicitado por : </b></td>
							<td colspan="3">
								<div class="ui-widget">
									<input type="text" id="med_derivador" name="med_derivador">					
									<a id ="but_display_med_deirv" tabindex="-1" class="ui-button ui-widget ui-state-default ui-button-icon-only custom-combobox-toggle ui-corner-right" role="button" title="Show All Items"><span class="ui-button-icon-primary ui-icon ui-icon-triangle-1-s"></span><span class="ui-button-text"></span></a>
									<input type="hidden" name="idmedicoderivador" id="idmedicoderivador">	
									<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('medicos_derivadores_02.asp?Tipo=A')"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agragar Medico Derivador"></a> 
								</div>						
							</td>					
					</tr>		
					<% 'if l_tipo = "M" then %>
					<tr>
					    <td align="right"><b>Precio:</b></td>
						<td colspan="3">
							<input align="right" type="text" name="precio" size="20" maxlength="20" value="0">
							<input type="hidden" name="precio2" value="">							
						</td>
					</tr>		
					<%' End If %>			
					<tr>
					    
						<td colspan="4">
							&nbsp;						
						</td>
					</tr>						

					<tr>
						<td  align="right" nowrap><b>Medio de Pago: </b></td>
						<td colspan="3"><select name="idmediodepago" size="1" style="width:200;" onchange="ctrolmetodopago();">
								<option value=0 selected>Seleccione un Medio</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM mediosdepago "
								l_sql = l_sql & " where empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY titulo "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("titulo") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idmediodepago.value="<%= l_idmediodepago %>"</script>

						</td>					
					</tr>		
					<tr>
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td colspan="3"><select name="idobrasocial" size="1" style="width:200;">
								<option value=0 selected>Seleccione una OS</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales "
								l_sql  = l_sql  & " WHERE isnull(obrassociales.flag_particular,0) = 0 "	
								l_sql = l_sql & " AND empnro = " & Session("empnro")								
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idobrasocial.value="<%= l_idobrasocial %>"</script>
							<script>ctrolmetodopago();</script>
						</td>					
					</tr>		
					<tr>
					    <td align="right"><b>Nro:</b></td>
						<td>
							<input   type="text" name="nro" size="20" maxlength="20" value="<%'= l_nro %>">
						</td>					
					</tr>		
					<tr>
					    <td align="right"><b>Importe:</b></td>
						<td>
							<input align="right" type="text" name="importe" size="20" maxlength="20" value="0">
							<input type="hidden" name="importe2" value="">
						</td>					
					</tr>								
					
					
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th">
		<a class=sidebtnABM href="Javascript:Validar_Formulario_visita()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida"  style="visibility=hidden;" src="" width="0%" height="0%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
		<!--	PARAMETRIZACION DE VENTANAS MODALES        -->
		<!--	URL´s        -->
		<input type="hidden" id="url_AM" value="medicos_derivadores_03.asp">
		<input type="hidden" id="url_valid_06" value="medicos_derivadores_06.asp">	
		<input type="hidden" id="url_baja" value="0">	
		<input type="hidden" id="url_base_baja" value="medicos_derivadores_04.asp?">	
		
		<!--	DIV´s Dialogos       -->
		<input type="hidden" id="id_dialog" value="dialog">
		<input type="hidden" id="width_dialog" value="300">
		<input type="hidden" id="height_dialog" value="auto">		
		<div id="dialog" title="Nuevo Medico Derivador"> 			</div>	  
		
		<input type="hidden" id="id_dialogAlert" value="dialogAlert">
		<div id="dialogAlert" title="Mensaje">				</div>	
		
		<input type="hidden" id="id_dialogConfirmDelete" value="dialogConfirmDelete">
		<div id="dialogConfirmDelete" title="Consulta">		</div>		
		<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->	
</body>
</html>
