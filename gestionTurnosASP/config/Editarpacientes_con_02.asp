
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
dim l_comentario
dim idrecursoreservable

Dim l_ventana

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_id = request.querystring("cabnro")

  Dim l_dnioblig
  Dim l_hc
  
  l_dnioblig  = request("dni")
  l_hc  = request("hc")

l_ventana = request.querystring("ventana")

'response.write l_tipo

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<!--<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Editar Pacientes</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>
<script>
function Validar_Formulario(){

var s = document.datos.osid;

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
<% If l_dnioblig = "S" then %>
if (document.datos.dni.value == ""){
	alert("Debe ingresar el DNI del Paciente.");
	document.datos.dni.focus();
	return;
}
if (isNaN(document.datos.dni.value)) {
	alert("El D.N.I. debe ser numerico.");
	document.datos.dni.focus();
	return;
}
<% End If %>
/*
if (document.datos.nrohistoriaclinica.value == ""){
	alert("Debe ingresar el Nro de Historia Clinica o ingresar 0.");
	document.datos.nrohistoriaclinica.focus();
	return;
}
if (isNaN(document.datos.nrohistoriaclinica.value)) {
	alert("El Nro de Historia Clinica debe ser numerico.");
	document.datos.nrohistoriaclinica.focus();
	return;
}
*/
/*
if (document.datos.domicilio.value == ""){
	alert("Debe ingresar el Domicilio del Paciente.");
	document.datos.domicilio.focus();
	return;
}
*/
if (document.datos.tel.value == ""){
	alert("Debe ingresar el Telefono del Paciente.");
	document.datos.tel.focus();
	return;
}
<% if l_hc = "S" then %>
if (document.datos.nrohistoriaclinica.value == ""){
	alert("Debe ingresar el Nro de Historia Clinica.");
	document.datos.nrohistoriaclinica.focus();
	return;
}
if (document.datos.nrohistoriaclinica.value == "0"){
	alert("Debe ingresar el Nro de Historia Clinica.");
	document.datos.nrohistoriaclinica.focus();
	return;
}
<% End If %>

// Texto seleccionado:  s.options[s.selectedIndex].text;
//alert(s.options[s.selectedIndex].text);
document.datos.os.value = s.options[s.selectedIndex].text;

var d=document.datos;
document.valida.location = "editarpacientes_con_06.asp?tipo=<%= l_tipo%>&id="+document.datos.id.value + "&dni="+document.datos.dni.value;

//valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	//document.datos.coudes.focus();
}


function EncontrePaciente(id, apellido, nombre, nrohistoriaclinica, dni, domicilio, tel){
	document.datos.pacienteid.value = id;
	document.datos.apellido.value = apellido;
	document.datos.nombre.value = nombre;
	document.datos.nrohistoriaclinica.value = nrohistoriaclinica;
	document.datos.dni.value = dni;
	document.datos.domicilio.value = domicilio;
	document.datos.tel.value = tel;
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

function Mayuscula(cadena){

	cadena.value = cadena.value.toUpperCase();
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
			l_idobrasocial  = "0"
			idrecursoreservable = ""
			l_nrohistoriaclinica = "0"
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM clientespacientes "
		'l_sql = l_sql & " INNER JOIN ser_servicio ON ser_servicio.sercod = ser_legajo.legpar1 "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
	    	l_apellido      = l_rs("apellido")
	    	l_nombre        = l_rs("nombre")
			l_nrohistoriaclinica = l_rs("nrohistoriaclinica")
	    	l_dni           = l_rs("dni")
	    	l_domicilio     = l_rs("domicilio")
			l_tel           = l_rs("telefono")
			if isnull(l_rs("idobrasocial")) then
				l_idobrasocial  = 0
			else
				l_idobrasocial  = l_rs("idobrasocial")
			end if
			'l_idrecursoreservable = l_rs("idrecursoreservable")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.apellido.focus();">
<form name="datos" action="Editarpacientes_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="hidden" name="id" value="<%= l_id %>">
<input type="hidden" name="pacienteid" value="<%'= l_id %>">
<input type="hidden" name="ventana" value="<%= l_ventana %>">
<input type="hidden" name="os" value="">

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
					    <td align="right"><b>Apellido (*):</b></td>
						<td>
							<input type="text" name="apellido" size="20" maxlength="20" onkeydown="Javascript:Mayuscula(this);" value="<%= l_apellido %>">							
						</td>
					    <td align="right"><b>Nombre (*):</b></td>						
						<td>
							<input type="text" name="nombre" size="20" maxlength="20" onkeydown="Javascript:Mayuscula(this);" value="<%= l_nombre %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I.:<% If l_dnioblig = "S" then %> (*)<% End If %></b></td>
						<td>
							<input type="text" name="dni" size="20" maxlength="20" value="<%= l_dni %>">
						</td>
					    <td align="right"><b>Tel&eacute;fono (*):</b></td>
						<td>
							<input type="text" name="tel" size="20" maxlength="20" value="<%= l_tel %>">
						</td>						
					
						<!--
					    <td align="right"><b>Nro. Historia Cl&iacute;nica:</b></td>
						<td>
							<input type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
						</td>	-->					
					</tr>
					<tr>
					    <td align="right"><b>Domicilio:</b></td>
						<td>
							<input type="text" name="domicilio" size="20" maxlength="20" value="<%= l_domicilio %>">
						</td>						

						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td ><select name="osid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una OS</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.osid.value="<%= l_idobrasocial %>"</script>
						</td>					
					</tr>
					<tr>
					    <td align="right"><b> Historia Cl&iacute;nica <% If l_hc = "S" then %> (*)<% End If %>:</b></td>
						<td>
							<input type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
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
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida"  src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
