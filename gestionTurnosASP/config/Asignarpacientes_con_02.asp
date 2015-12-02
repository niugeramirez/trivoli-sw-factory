
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_clientepacienteid
dim l_apellido
dim l_nombre  
dim l_nrohistoriaclinica
dim l_dni     
dim l_tel
dim l_domicilio
dim l_idobrasocial
dim l_descripcion
dim l_comentario
dim l_idmedicoderivador
dim l_med_derivador
Dim l_idpractica
Dim l_ventana
Dim l_iduser
Dim l_agenda

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_agenda = request.querystring("agenda")
l_tipo = request.querystring("tipo")
l_id = request.querystring("cabnro")


if l_tipo = "A" then
	l_ventana = 1
else
	l_ventana = 2
end if

'response.write l_tipo

%>
<html>
<head>
<title>Asignar Pacientes</title>

<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css" />
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery-ui.js"></script>
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


</head>


<script>
function Validar_Formulario_turno(){

if (document.datos.apellido.value == ""){
	alert("Debe ingresar el Apellido del Paciente.");
	document.datos.apellido.focus();
	return;
}

if (document.datos.nombre.value == ""){
	alert("Debe ingresar el Nombre del Paciente.");
	document.datos.nombre.focus();
	return;
}

if (document.datos.tel.value == ""){
	alert("Debe ingresar el Telefono del Paciente.");
	document.datos.tel.focus();
	return;
}

if (document.datos.practicaid.value == "0"){
	alert("Debe ingresar una Practica.");
	document.datos.practicaid.focus();
	return;
}

if (document.datos.iduser.value == "0"){
	alert("Debe ingresar quien Ingresa el Turno.");
	document.datos.iduser.focus();
	return;
}

valido_turno();
}

function valido_turno(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	//document.datos.coudes.focus();
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
	//document.datos.coudes.focus();
}



function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function BuscarPaciente(){
	abrirVentana('Buscarpacientes_con_00.asp?Tipo=A&Alta=S&dni=N&hc=N','',600,250);
}

function Editar(){ 

	if (document.datos.pacienteid.value == 0){
		alert("Debe ingresar el Paciente.");
		document.datos.pacienteid.focus();
		return;
	}; 
	
	parent.abrirVentana('Editarpacientes_con_02.asp?Tipo=M&Ventana=<%= l_ventana %>&cabnro='+document.datos.pacienteid.value ,'',600,250) 
}


</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
 	    	l_apellido      = ""
	    	l_nombre        = ""
	    	l_dni           = ""
	    	l_domicilio     = ""
			l_tel           = ""
			l_idobrasocial  = ""
			l_idmedicoderivador = "0"
			l_clientepacienteid = "0"
			
			l_sql = "SELECT  * "
			l_sql = l_sql & " FROM config "
			l_sql  = l_sql  & " WHERE config.empnro = " & Session("empnro")
			
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
		    	l_idpractica = l_rs("idpractica")
		    else
				l_idpractica = "0"
			end if
			l_rs.Close			
			l_iduser  = session("loguinUser")
			
			
	Case "M":
		
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  turnos.id id , turnos.idpractica, turnos.idmedicoderivador, turnos.comentario, turnos.iduseringresoturno "
		l_sql = l_sql & " , clientespacientes.id clientepacienteid, clientespacientes.apellido, clientespacientes.nombre , clientespacientes.nrohistoriaclinica "
		l_sql = l_sql & " , clientespacientes.dni , clientespacientes.domicilio, clientespacientes.telefono , clientespacientes.idobrasocial "
		l_sql = l_sql & " , obrassociales.id idobrasocial , obrassociales.descripcion , medicos_derivadores.nombre as med_derivador "
		l_sql = l_sql & " FROM turnos "
		l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente" 
		l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
		l_sql = l_sql & " LEFT JOIN medicos_derivadores ON medicos_derivadores.id = turnos.idmedicoderivador "
		l_sql  = l_sql  & " WHERE turnos.id = " & l_id
		'response.write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_id            = l_rs("id")
			l_clientepacienteid = l_rs("clientepacienteid")
	    	l_apellido      = l_rs("apellido")
	    	l_nombre        = l_rs("nombre")
			l_nrohistoriaclinica = l_rs("nrohistoriaclinica")
	    	l_dni           = l_rs("dni")
	    	l_domicilio     = l_rs("domicilio")
			l_tel           = l_rs("telefono")
			l_idobrasocial  = l_rs("idobrasocial")
			l_descripcion   = l_rs("descripcion")
			l_idpractica    = l_rs("idpractica")
			l_idmedicoderivador = l_rs("idmedicoderivador")
			l_med_derivador = l_rs("med_derivador")
			'response.write l_idmedicoderivador
			l_comentario    = l_rs("comentario")
			l_iduser        = l_rs("iduseringresoturno")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" <% if l_tipo <> "M" then %> onload="javascript:BuscarPaciente();" <% End If %>>
<form name="datos" action="Asignarpacientes_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="id" value="<%= l_id %>">
<input type="hidden" name="pacienteid" value="<%= l_clientepacienteid %>">

<input type="hidden" name="osid" value="<%= l_idobrasocial %>">
<input type="hidden" name="agenda" value="<%= l_agenda %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>&nbsp;</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
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
						<td colspan="2" align="left"><a href="Javascript:Editar();"><img src="/turnos/shared/images/Modificar_16.png" border="0" title="Editar Paciente"></a></td>
					    					
					</tr>				
					
					<tr>
						<td>&nbsp;
						</td>
					</tr>
						
					<tr>
						<td  align="right" nowrap><b>Practica (*): </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
								l_sql  = l_sql  & " WHERE practicas.empnro = " & Session("empnro")
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.practicaid.value="<%= l_idpractica %>"</script>
						</td>					
					</tr>	
					<tr>
					    <td align="right"><b>Comentario:</b></td>
						<td colspan="3">
							<input type="text" name="comentario" size="72" maxlength="100" value="<%= l_comentario %>">
						</td>
					   				
					</tr>						
					<tr>
						<td  align="right" nowrap><b>Solicitado por : </b></td>
						<td colspan="3">
							<div class="ui-widget">
								<input type="text" id="med_derivador" name="med_derivador">					
								<a id ="but_display_med_deirv" tabindex="-1" class="ui-button ui-widget ui-state-default ui-button-icon-only custom-combobox-toggle ui-corner-right" role="button" title="Show All Items"><span class="ui-button-icon-primary ui-icon ui-icon-triangle-1-s"></span><span class="ui-button-text"></span></a>
								<input type="hidden" name="idmedicoderivador" id="idmedicoderivador">
								<script>document.datos.idmedicoderivador.value="<%= l_idmedicoderivador %>"</script>	
								<script>document.datos.med_derivador.value="<%= l_med_derivador %>"</script>	
								<a id="abrirAlta" class="sidebtnABM" href="Javascript:abrirDialogo('medicos_derivadores_02.asp?Tipo=A')"><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agragar Medico Derivador"></a> 
							</div>
						</td>					
					</tr>			
					
					<tr> 
						<td  align="right" nowrap><b>Ingresado por : </b></td>
						<td colspan="3"><select name="iduser" size="1" style="width:200;">
								<option value=0 selected>Seleccionar un Usuario</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM user_per "
								l_sql  = l_sql  & " WHERE iduser <> 'sa' "	
								l_sql  = l_sql  & " AND user_per.empnro = " & Session("empnro")								
								l_sql  = l_sql  & " ORDER BY iduser "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("iduser") %> > 
								<%= l_rs("usrnombre") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.iduser.value="<%= l_iduser %>"</script> 
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
		<a class=sidebtnABM href="Javascript:Validar_Formulario_turno()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="" width="0%" height="0%"></iframe> 
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
