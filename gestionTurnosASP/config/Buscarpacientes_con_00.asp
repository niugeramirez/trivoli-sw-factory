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
  
  Dim l_alta
  Dim l_dni
  Dim l_hc
  
  l_alta  = request("Alta")
  l_dni  = request("dni")
  l_hc  = request("hc")
  
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Buscar Pacientes</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script>

 parent.document.datos.apellido.value = 'pepe';

function orden(pag){
	abrirVentana('../shared/asp/orden_browse.asp?pagina='+pag+'&lista=<%= l_orden %>&campos=<%= l_camposOr%>&filtro='+escape(document.ifrm.datos.filtro.value),'',350,160)
}

function filtro(pag){
	abrirVentana('../shared/asp/filtro_browse.asp?pagina='+pag+'&campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&orden='+document.ifrm.datos.orden.value,'',250,160);
}

function llamadaexcel(){ 
	if (filtro == "")
		Filtro(true);
	else
		abrirVentana("contracts_con_excel.asp?orden=" + document.ifrm.datos.orden.value + "&filtro=" + escape(document.ifrm.datos.filtro.value),'execl',250,150);
}


function Buscar(){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";

	// fec. desde
	/*
	if (document.datos.fechadesde.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechadesde)){
			document.datos.filtro.value += " ser_legajo.legfecing >= " + cambiafecha(document.datos.fechadesde.value,true,1) + "";
			tieneotro = "si";
		}else{
			estado = "no";
		}
	}
	// fec. hasta
	if (document.datos.fechahasta.value != ""){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND " ;
			tieneotro = "si";
		}
		if (validarfecha(document.datos.fechahasta)){
			document.datos.filtro.value += " ser_legajo.legfecing <= " + cambiafecha(document.datos.fechahasta.value,true,1) + "";
		}else{
			estado = "no";
		}
	}
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
		estado = "no";
	}	
	
	// Servicio Local
	if (document.datos.sernro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.legpar1 = '" + document.datos.sernro.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.legpar1 = '" + document.datos.sernro.value + "'";
		}
		tieneotro = "si";
	}	
	// Derecho Vulnerado
	if (document.datos.pronro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.pronro = '" + document.datos.pronro.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.pronro = '" + document.datos.pronro.value + "'";
		}
		tieneotro = "si";
	}		*/
	// Apellido
	if (document.datos.legape.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND clientespacientes.apellido like '*" + document.datos.legape.value + "*'";
		}else{
			document.datos.filtro.value += " clientespacientes.apellido like '*" + document.datos.legape.value + "*'";
		}
		tieneotro = "si";
	}		
	// Nombre
	if (document.datos.legnom.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND clientespacientes.nombre like '*" + document.datos.legnom.value + "*'";
		}else{
			document.datos.filtro.value += " clientespacientes.nombre like '*" + document.datos.legnom.value + "*'";
		}
		tieneotro = "si";
	}		
	// Nro. Historia Clinica
	if (document.datos.nrohistoriaclinica.value != 0){
		if (tieneotro == "si"){
			//document.datos.filtro.value += " AND clientespacientes.nrohistoriaclinica = '" + document.datos.nrohistoriaclinica.value + "'";
			document.datos.filtro.value += " AND clientespacientes.nrohistoriaclinica like '*" + document.datos.nrohistoriaclinica.value + "*'";
		}else{
			//document.datos.filtro.value += " clientespacientes.nrohistoriaclinica = '" + document.datos.nrohistoriaclinica.value + "'";
			document.datos.filtro.value += " clientespacientes.nrohistoriaclinica like '*" + document.datos.nrohistoriaclinica.value + "*'";
		}
		tieneotro = "si";
	}				
	// DNI
	if (document.datos.legdni.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND clientespacientes.dni like '*" + document.datos.legdni.value + "*'";
		}else{
			document.datos.filtro.value += " clientespacientes.dni like '*" + document.datos.legdni.value + "*'";
		}
		tieneotro = "si";
	}					
	// Domicilio
	/*if (document.datos.legdom.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND clientespacientes.domicilio like '*" + document.datos.legdom.value + "*'";
		}else{
			document.datos.filtro.value += " clientespacientes.domicilio like '*" + document.datos.legdom.value + "*'";
		}
		tieneotro = "si";
	}		
	// Medida de Proteccion
	if (document.datos.mednro.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND ser_legajo.mednro = '" + document.datos.mednro.value + "'";
		}else{
			document.datos.filtro.value += " ser_legajo.mednro = '" + document.datos.mednro.value + "'";
		}
		tieneotro = "si";
	}		*/					

	
	if (Trim(document.datos.filtro.value) == ""){
		alert("Debe ingresar el Filtro.");
		document.datos.legape.focus();
		return;
	}
	
	if (estado == "si"){
		window.ifrm.location = 'Buscarpacientes_con_01.asp?asistente=0&filtro=' + document.datos.filtro.value;
	}
}


function VolverdelAltaPaciente(ape){

	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";

	
	document.datos.legape.value = ape;
	
	// Apellido
	if (document.datos.legape.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND clientespacientes.apellido like '*" + document.datos.legape.value + "*'";
		}else{
			document.datos.filtro.value += " clientespacientes.apellido like '*" + document.datos.legape.value + "*'";
		}
		tieneotro = "si";
	}		
		

	//alert(document.datos.filtro.value);
	
	if (estado == "si"){
		window.ifrm.location = 'Buscarpacientes_con_01.asp?asistente=0&filtro=' + document.datos.filtro.value;
	}
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

function AltaPaciente(){
	abrirVentana('Editarpacientes_con_02.asp?Tipo=A&ventana=3&dni=<%= l_dni %>&hc=<%= l_hc %>','',600,250);
}

function Limpiar(){

	//document.datos.fechadesde.value = "<%= date() - 1 %>";
	//document.datos.fechahasta.value = "<%= date() %>";

	//document.datos.sernro.value     = 0;
	//document.datos.pronro.value     = 0;	
/*	
	document.datos.pronro.value     = 0;
	document.datos.ternro.value     = 0;
	document.datos.pornro.value     = 0;
	document.datos.clinro.value     = 0;	
	document.datos.ctrnum.value     = 0;	
	document.datos.txtctrnum.value  = "";	
*/	
	window.ifrm.location = 'pacientes_con_01.asp';
}


function fnctrnum(valor){
	if (valor == 0){
		document.datos.txtctrnum.disabled = true;
		document.datos.txtctrnum.className = "deshabinp";
	}else{
		document.datos.txtctrnum.disabled = false;
		document.datos.txtctrnum.className = "habinp";
	}
}



</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.legape.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">
			<% if l_alta = "S" then %>
			<a href="Javascript:AltaPaciente();"><img src="/turnos/shared/images/Agregar_24.png" border="0" title="Alta Paciente"></a>	
			<% End If %>
		  	  
		  <!--
		  <a class=sidebtn href="Javascript:orden('../../config/contracts_con_01.asp');">Orden</a>
		  -->
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		  
		  </td>
        </tr>		
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<form name="datos">
					<input type="hidden" name="filtro" value="">
					<!--
					<tr>
						<td align="right"><b>Fec. Desde: </b></td>
						<td><input  type="text" name="fechadesde" size="10" maxlength="10" value="<%'= Date() - 1 %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Fec. Hasta: </b></td>
						<td><input  type="text" name="fechahasta" size="10" maxlength="10" value="<%'= Date() %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta);"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
					</tr>
					
					<tr>
						<td align="right"><b>Fec. Ingreso Desde: </b></td>
						<td><input  type="text" name="fechadesde" size="10" maxlength="10" value="<%'= Date() - 1 %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Fec. Ingreso Hasta: </b></td>
						<td><input  type="text" name="fechahasta" size="10" maxlength="10" value="<%'= Date() %>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta);"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
					</tr>-->
					<!-- 
					<tr>
						<td  align="right" nowrap><b>Servicio Local: </b></td>
						<td><select name="sernro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM ser_servicio "
								'l_sql  = l_sql  & " ORDER BY sercod "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("sercod") %> > 
								<%'= l_rs("serdes") %> (<%'=l_rs("sercod")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.sernro.value= "0"</script>
						</td>					
						<td  align="right" nowrap><b>Derecho Vulnerado: </b></td>
						<td><select name="pronro" size="1" style="width:150;">
								<option value=0 selected>Todos</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM ser_problematica "
								'l_sql  = l_sql  & " ORDER BY prodes "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("pronro") %> > 
								<%'= l_rs("prodes") %> (<%'=l_rs("pronro")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.pronro.value= "0"</script>
						</td>
					</tr>  -->
					<tr>
						<td align="right"><b>Apellido: </b></td>
						<td><input  type="text" name="legape" size="21" maxlength="21" value="" >
						</td>
						<td align="right"><b>Nombre: </b></td>
						<td><input  type="text" name="legnom" size="21" maxlength="21" value="" >
						</td>					

					</tr>
					<tr>
						<td align="right"><b>D.N.I.: </b></td>
						<td><input  type="text" name="legdni" size="21" maxlength="21" value="" >
						</td>		
					    <td align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
						<td>
							<input type="text" name="nrohistoriaclinica" size="21" value="">
						</td>											
					</tr>			
							


					<tr>
						<td align="right" colspan="2" >
						<table border="0" cellpadding="0" cellspacing="0" bgcolor="Navy">
								<tr>
								
										<!--<td ><img src="../shared/images/gen_rep/boton_01.gif" width="5.9"></td>-->
										<td ><a class="sidebtnABM" href="Javascript:Buscar();"><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar"></a></td>
										<!--<td  background="../shared/images/gen_rep/boton_05.gif"><img src="../shared/images/gen_rep/boton_03.gif" height="15"></td>-->
										<td ><a class="sidebtnABM" href="Javascript:Limpiar();"><img  src="/turnos/shared/images/Limpiar_24.png" border="0" title="Limpiar"></a></td>
										<!--<td ><img src="../shared/images/gen_rep/boton_06.gif"></td>-->
								</tr>
							</table>
						</td>

						
												
					</tr>					

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
			</form>		
      </table>
</body>

<script>
	//Buscar();
</script>
</html>
