
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_apellido
dim l_nombre  
dim l_nrohistoriaclinica
dim l_dni     
dim l_tel
dim l_domicilio
dim l_idobrasocial
dim l_descripcion
dim l_comentario
dim l_idrecursoreservable
Dim l_idpractica
'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_id = request.querystring("cabnro")

'response.write l_tipo

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Asignar Pacientes</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){




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

if (document.datos.dni.value == ""){
	alert("Debe ingresar el DNI del Paciente.");
	document.datos.dni.focus();
	return;
}

if (document.datos.domicilio.value == ""){
	alert("Debe ingresar el Domicilio del Paciente.");
	document.datos.domicilio.focus();
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

/*
if (document.datos.tipopenro.value == 0){
	alert("Debe ingresar el Tipo de Operación.");
	document.datos.tipopenro.focus();
	return;
}
if (document.datos.tipbuqnro.value == 0){
	alert("Debe ingresar el Tipo de Buque.");
	document.datos.tipbuqnro.focus();
	return;
}
if (document.datos.agenro.value == 0){
	alert("Debe ingresar la Agencia.");
	document.datos.agenro.focus();
	return;
}

if ((document.datos.buqfecdes.value != "")&&(!validarfecha(document.datos.buqfecdes))){
	 document.datos.buqfecdes.focus();
	 return;
}

if ((document.datos.buqfechas.value != "")&&(!validarfecha(document.datos.buqfechas))){
	 document.datos.buqfechas.focus();
	 return;
}

if ((document.datos.buqfecdes.value != "")&&(document.datos.buqfechas.value != "") ){

	if (!(menorque(document.datos.buqfecdes.value,document.datos.buqfechas.value))) {
			alert("La Fecha de Comienzo debe ser menor o igual que la Fecha de Termino.");
			document.datos.buqfecdes.focus();
		    return;			
	}		
}	
*/

/*
var d=document.datos;
document.valida.location = "pacientes_con_06.asp?tipo=<%= l_tipo%>&counro="+document.datos.counro.value + "&coudes="+document.datos.coudes.value;
*/
valido();
}

function valido(){
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
	abrirVentana('Buscarpacientes_con_00.asp?Tipo=A','',600,250);
}

</script>
<% 
select Case l_tipo
	Case "A":
 	    	l_apellido      = ""
	    	l_nombre        = ""
	    	l_dni           = ""
	    	l_domicilio     = ""
			l_tel           = ""
			l_idobrasocial  = ""
			l_idrecursoreservable = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  turnos.*, clientespacientes.id clientepacienteid, clientespacientes.* , obrassociales.id idobrasocial , obrassociales.descripcion"
		l_sql = l_sql & " FROM turnos "
		l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente" 
		l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
		l_sql  = l_sql  & " WHERE turnos.id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
	    	l_apellido      = l_rs("apellido")
	    	l_nombre        = l_rs("nombre")
			l_nrohistoriaclinica = l_rs("nrohistoriaclinica")
	    	l_dni           = l_rs("dni")
	    	l_domicilio     = l_rs("domicilio")
			l_tel           = l_rs("telefono")
			l_idobrasocial  = l_rs("idobrasocial")
			l_descripcion   = l_rs("descripcion")
			l_idpractica    = l_rs("idpractica")
			l_idrecursoreservable = l_rs("idrecursoreservable")
			l_comentario = l_rs("comentario")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" <% if l_tipo <> "M" then %> onload="javascript:BuscarPaciente();" <% End If %>>
<form name="datos" action="Asignarpacientes_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="id" value="<%= l_id %>">
<input type="hidden" name="pacienteid" value="<%'= l_id %>">

<input type="hidden" name="osid" value="<%= l_idobrasocial %>">

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
					<a href="Javascript:BuscarPaciente();"><img src="/turnos/shared/images/BuscarPaciente.png" border="0" alt="Buscar Paciente"></a>		
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
					    					
					</tr>				
					
					<!--
					<tr>
					    <td align="right" ><b>Fec. Nac.:</b></td>
						<td align="left"  >
						    <input type="text" name="legfecnac" size="10" maxlength="10" value="<%'= l_legfecnac %>">
							<a href="Javascript:Ayuda_Fecha(document.datos.legfecnac)"><img src="/turnos/shared/images/cal.gif" border="0"></a>
						</td>
						<td align="right"><b>Teléfono:</b></td>
						<td>
							<input type="text" name="legtel" size="20" maxlength="20" value="<%'= l_legtel %>">
						</td>						
					</tr>
					
					
					<tr>
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td colspan="3"><select class="deshabinp"  name="osid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una OS</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM obrassociales "
								'l_sql  = l_sql  & " ORDER BY descripcion "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("id") %> > 
								<%'= l_rs("descripcion") %> </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.osid.value=0 "<%'= l_pronro %>"</script>
						</td>					
					</tr>
					-->
					<tr>
						<td>&nbsp;
						</td>
					</tr>
					<tr>
						<td>&nbsp;
						</td>
					</tr>										
					<tr>
						<td  align="right" nowrap><b>Practica: </b></td>
						<td colspan="3"><select name="practicaid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una Practica</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM practicas "
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
						<td colspan="3"><select name="idrecursoreservable" size="1" style="width:200;">
								<option value=0 selected>Seleccione un Recurso</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.idrecursoreservable.value="<%= l_idrecursoreservable %>"</script>
						</td>					
					</tr>					
					
									<!-- 
					<tr>
					    <td align="right"><b>Madre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenommad" size="20" maxlength="20" value="<%'= l_legapenommad %>">
						</td>
						<td align="right"><b>Dom:</b></td>						
						<td>
							<input type="text" name="legdommad" size="20" maxlength="20" value="<%'= l_legdommad %>">
							<b>Tel:</b> <input type="text" name="legtelmad" size="10" maxlength="20" value="<%'= l_legtelmad %>">						
						</td>							
					</tr>																				
					<tr>
					    <td align="right"><b>Padre - Apellido y Nombre:</b></td>
						<td>
							<input type="text" name="legapenompad" size="20" maxlength="20" value="<%'= l_legapenompad  %>">
						</td>
						<td align="right"><b>Dom:</b></td>												
						<td>
							<input type="text" name="legdompad" size="20" maxlength="20" value="<%'= l_legdompad %>">
							<b>Tel:</b> <input type="text" name="legtelpad" size="10" maxlength="20" value="<%'= l_legtelpad %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>Instituciones Intervinientes:</b></td>
						<td colspan="3">
							<input type="text" name="legins" size="80" maxlength="20" value="<%'= l_legins %>">
						</td>
					</tr>																				
					<tr>
					    <td align="right"><b>Instituciones Educativas:</b></td>
						<td colspan="3">
							<input type="text" name="leginsedu" size="80" maxlength="20" value="<%'= l_leginsedu %>">
						</td>
					</tr>																									
					<tr>
					    <td align="right"><b>Cobertura Social de la Familia:</b></td>
						<td colspan="3">
							<input type="text" name="legcobsoc" size="80" maxlength="20" value="<%'= l_legcobsoc %>">
						</td>
					</tr>																														
					<tr>
					    <td align="right"><b>Estrategias de Intervención:</b></td>
						<td colspan="3">
							<input type="text" name="legabo" size="80" maxlength="20" value="<%'= l_legabo %>">
						</td>
					</tr>					
					<tr>
						<td align="right"><b>Medidas Protección:</b></td>
						<td colspan="3"><select name="mednro" size="1" style="width:150;">
								<option value=0 selected>&nbsp;</option>
								<%'Set l_rs = Server.CreateObject("ADODB.RecordSet")
								'l_sql = "SELECT  * "
								'l_sql  = l_sql  & " FROM ser_medida "
								'l_sql  = l_sql  & " ORDER BY meddes "
								'rsOpen l_rs, cn, l_sql, 0
								'do until l_rs.eof		%>	
								<option value= <%'= l_rs("mednro") %> > 
								<%'= l_rs("meddes") %> (<%'=l_rs("mednro")%>) </option>
								<%'	l_rs.Movenext
								'loop
								'l_rs.Close %>
							</select>
							<script>document.datos.mednro.value= "<%'= l_mednro %>"</script>
						</td>					
					</tr>					
					 -->						
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
