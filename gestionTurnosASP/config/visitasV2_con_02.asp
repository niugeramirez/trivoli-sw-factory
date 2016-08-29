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

dim l_visitaid
dim l_pacienteid
dim l_apellido
dim l_nombre
dim l_dni
dim l_nrohistoriaclinica
dim l_tel
dim l_domicilio
dim l_idobrasocial
dim l_osnombre
dim l_idmediodepago

dim l_practicasrealizadasid
dim l_idpractica 
dim l_idsolicitadapor
dim l_med_derivador
dim l_precio

l_tipo = request.querystring("tipo")
l_idrecursoreservable = request.querystring("idrecursoreservable")
l_calfec  = request.querystring("fechadesde")
l_visitaid = request.querystring("visitaid")
l_practicasrealizadasid = request.querystring("practicasrealizadasid")


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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


<!--Inicio ventanas modales-->
<script>
//Esto va antes de la importacion de ventanas_modales_custom.js
function Validaciones_locales_med_der(){
	//devuelvo siempre verdadero en este caso de modo de hacer los controles en el 06
	return true;
}
</script>
<script>


function BuscarPaciente(){
	
	abrirDialogo('dialog_cont_BusqPac','BuscarpacientesV2_00.asp?Tipo=A&Alta=S&fn_asign_pac=volver_AsignarPaciente&dnioblig=S&hcoblig=S',900,250);
}

$(document).ready(function() { 
								inicializar_dialogoContenedor(	"dialog_cont_BusqPac" 										//id_dialog
																); 																	
								inicializar_dialogoABM(	"dialogAVST" 										//id_dialog
														,"medicos_derivadores_06.asp"							//url_valid_06
														,"medicos_derivadores_03.asp"							//url_AM
														,"dialogAlertVisitas"									//id_dialogAlert	
														,"datos_med_der"										//id_form_datos		
														,null //window.parent.ifrm.location					//location_reload
														,Validaciones_locales_med_der							//funcion_Validaciones_locales	
														,null //"ifrm"											//id_ifrm_form_datos														
														); 
								<% if l_tipo <> "M" then %>
								BuscarPaciente();
								<% End If %>																		
																						
							});
</script>
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

'obtengo la Obra Social Particular
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
		l_pacienteid = 0
		l_visitaid = 0

		l_idpractica = 0
		l_idsolicitadapor = 0
		l_precio = 0		
	Case "M":				

		l_sql = "SELECT  clientespacientes.id l_pacienteid,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.dni, clientespacientes.nrohistoriaclinica nrohistoriaclinica , isnull(clientespacientes.idobrasocial,0) idobrasocial, obrassociales.descripcion osnombre"
		l_sql = l_sql & " ,clientespacientes.telefono, clientespacientes.domicilio, visitas.fecha "
		l_sql = l_sql & " FROM visitas "
		l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
		l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
		l_sql = l_sql & " WHERE visitas.id = "&l_visitaid
	
		
		'response.write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then					
			l_pacienteid = l_rs("clientepacienteid")			
			l_apellido      = l_rs("apellido")
			l_nombre        = l_rs("nombre")
			l_nrohistoriaclinica = l_rs("nrohistoriaclinica")
			l_dni           = l_rs("dni")
			l_domicilio     = l_rs("domicilio")
			l_tel           = l_rs("telefono")			
			l_idobrasocial  = l_rs("idobrasocial")
			l_osnombre   = l_rs("osnombre")		
			l_calfec		=	l_rs("fecha")
			'response.write	"l_idobrasocial "&l_idobrasocial		
		end if
		l_rs.Close
		
		'response.write	"l_practicasrealizadasid "&l_practicasrealizadasid
		if l_practicasrealizadasid = "" or l_practicasrealizadasid = 0 then
			l_idpractica = 0
			l_idsolicitadapor = 0
			l_precio = 0	
		else
			l_sql = "SELECT practicasrealizadas.idpractica ,practicasrealizadas.idsolicitadapor ,practicasrealizadas.precio "
			l_sql = l_sql & " , medicos_derivadores.nombre as med_derivador "
			l_sql = l_sql & " FROM practicasrealizadas "
			l_sql = l_sql & " LEFT JOIN medicos_derivadores ON medicos_derivadores.id = practicasrealizadas.idsolicitadapor "
			l_sql  = l_sql  & " WHERE practicasrealizadas.id = " & l_practicasrealizadasid
			
			
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
				l_idpractica = l_rs("idpractica")
				l_idsolicitadapor = l_rs("idsolicitadapor") 
				l_precio = l_rs("precio")
				l_med_derivador = l_rs("med_derivador")
			end if
			l_rs.Close				
		end if
	
end select

if l_idobrasocial = l_osparticular then
	l_idmediodepago = 0
	l_idobrasocial = l_osparticular
else
	l_idmediodepago = 	l_mediodepagoos
end if
%>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0"   >
<form name="datosAVST" id="datosAVST" action="Submit_Formulario_visit();" onkeypress=""  target="valida">
	<input type="hidden" name="idrecursoreservable" value="<%= l_idrecursoreservable %>">
	<input type="hidden" name="visitaid" value="<%= l_visitaid %>">
	<input type="hidden" name="practicasrealizadasid" value="<%= l_practicasrealizadasid %>">
	
	<input type="hidden" name="pacienteid" value="<%= l_pacienteid %>">
	<input type="hidden" name="osid" value="<%= l_idobrasocial %>">
	<input type="hidden" name="calfec" value="<%= l_calfec %>">
	<input type="hidden" name="mediodepagoos" value="<%= l_mediodepagoos %>">
	<input type="hidden" name="osparticular" value="<%= l_osparticular %>">


	<table cellspacing="0" cellpadding="0" border="0">						
		<tr>
			<td align="right"><b>Apellido:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="apellido" size="20" maxlength="20" value="<%= l_apellido %>">							
			</td>
			<td colspan="3" align="right"><b>Nombre:</b></td>						
			<td>
				<input class="deshabinp" readonly="" type="text" name="nombre" size="20" maxlength="20" value="<%= l_nombre %>">								
			</td>
			<td colspan="4" align="left">
				<% if l_tipo <> "M" then %>
				<a href="Javascript:BuscarPaciente();"><img src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar Paciente"></a>		
				<% End If %>						
			</td>
		</tr>					
		<tr>
			<td align="right"><b>D.N.I.:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="dni" size="20" maxlength="20" value="<%= l_dni %>">
			</td>
			<td colspan="3" align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
			</td>	
			<td colspan="6" align="left">&nbsp;</td>			
		</tr>
		<tr>
			<td align="right"><b>Tel&eacute;fono:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="tel" size="20" maxlength="20" value="<%= l_tel %>">
			</td>
			<td colspan="3" align="right"><b>Domicilio:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="domicilio" size="20" maxlength="20" value="<%= l_domicilio %>">
			</td>	
			<td colspan="4" align="left">&nbsp;</td>			
		</tr>

		<tr>
			<td align="right"><b>Obra Social:</b></td>
			<td>
				<input class="deshabinp" readonly="" type="text" name="os" size="20" maxlength="20" value="<%= l_osnombre %>">
			</td>
			<td colspan="6" align="left">&nbsp;</td>
								
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
				<script>document.datosAVST.practicaid.value="<%= l_idpractica %>"</script>
			</td>
			<!-- PAGO -->
			<% if l_practicasrealizadasid = "" or l_practicasrealizadasid = 0 then%>
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
				<script>document.datosAVST.idmediodepago.value="<%= l_idmediodepago %>"</script>
			</td>
			<% end if %>
			<!-- PAGO -->							
		</tr>	
		
		<tr>
			<td  align="right" nowrap><b>Solicitado por : </b></td>
			<td colspan="3">
				<div class="ui-widget">
					<input type="text" id="med_derivador" name="med_derivador">					
					<a id ="but_display_med_deirv" tabindex="-1" class="ui-button ui-widget ui-state-default ui-button-icon-only custom-combobox-toggle ui-corner-right" role="button" title="Show All Items"><span class="ui-button-icon-primary ui-icon ui-icon-triangle-1-s"></span><span class="ui-button-text"></span></a>
					<input type="hidden" name="idmedicoderivador" id="idmedicoderivador">	
					<script>document.datosAVST.idmedicoderivador.value="<%= l_idsolicitadapor %>"</script>	
					<script>document.datosAVST.med_derivador.value="<%= l_med_derivador %>"</script>										
					<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('dialogAVST','medicos_derivadores_02.asp?Tipo=A',300,'auto')"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agragar Medico Derivador"></a> 
				</div>						
			</td>
			<!-- PAGO -->
			<% if l_practicasrealizadasid = "" or l_practicasrealizadasid = 0 then%>							
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
				<script>document.datosAVST.idobrasocial.value="<%= l_idobrasocial %>"</script>
				<script>ctrolmetodopago();</script>
			</td>
			<td align="right"><b>Nro:</b></td>
			<td>
				<input   type="text" name="nro" size="20" maxlength="20" value="">
			</td>	
			<% end if %>							
			<!-- PAGO -->							
		</tr>		
		
		<tr>
			<td align="right"><b>Precio:</b></td>
			<td colspan="3">
				<input align="right" type="text" name="precio" size="20" maxlength="20" value="<%= l_precio %>">
				<input type="hidden" name="precio2" value="">							
			</td>
			<!-- PAGO -->	
			<% if l_practicasrealizadasid = "" or l_practicasrealizadasid = 0 then%>							
			<td align="right"><b>Importe:</b></td>
			<td>
				<input align="right" type="text" name="importe" size="20" maxlength="20" value="0">
				<input type="hidden" name="importe2" value="">
			</td>	
			<% end if %>
			<!-- PAGO -->
		</tr>		
	</table>

	<iframe name="valida"  style="visibility=hidden;" src="" width="0%" height="0%"></iframe> 
</form>

<!--	PARAMETRIZACION DE VENTANAS MODALES        -->

<!--	DIV´s Dialogos       -->	
<div id="dialogAVST" title="Nuevo Medico Derivador"> 			</div>	  
	

<div id="dialog_cont_BusqPac" title="Buscar Pacientes">		</div>			
<!--	FIN DE PARAMETRIZACION DE VENTANAS MODALES -->	
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>		
</body>
</html>
