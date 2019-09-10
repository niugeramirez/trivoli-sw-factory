
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/pacientes_util.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_apellido
dim l_nombre  
dim l_nrohistoriaclinica
dim l_dni     
dim l_domicilio
dim l_telefono
dim l_idobrasocial

dim l_fecha_ingreso
Dim l_fechanacimiento
dim l_nro_obra_social
Dim l_sexo
Dim l_ciudad 

dim l_observaciones

dim l_oblig

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

'response.write l_tipo

%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Pacientes</title>
</head>
<script src="/turnos/shared/js/fn_valida.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_numeros.js"></script>

<!-- Comienzo Datepicker -->
<link rel="stylesheet" href="../js/themes/smoothness/jquery-ui.css">
<script src="../js/jquery-1.8.0.js"></script>
<script src="../js/jquery-ui.js"></script>  
<script src="../js/jquery.ui.datepicker-es.js"></script>
<script>
$(function () {
$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});

		
$( "#fecha_ingreso" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fechanacimiento" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->

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
if (isNaN(document.datos.dni.value)) {
	alert("El D.N.I. debe ser numerico.");
	document.datos.dni.focus();
	return;
}

if ((document.datos.nrohistoriaclinica.value == "" || document.datos.nrohistoriaclinica.value == 0)
	&& document.datos.gen_hist_num.checked == false ){
	alert("Debe ingresar el Nro de Historia Clinica o seleccionar la opcion para generar un nuevo numero.");
	document.datos.nrohistoriaclinica.focus();
	return;
}

if (isNaN(document.datos.nrohistoriaclinica.value)) {
	alert("El Nro de Historia Clinica debe ser numerico.");
	document.datos.nrohistoriaclinica.focus();
	return;
}

if (document.datos.telefono.value == ""){
	alert("Debe ingresar el Telefono del Paciente.");
	document.datos.telefono.focus();
	return;
}
if (document.datos.osid.value == "0"){
	alert("Debe ingresar la Obra Social del Paciente.");
	document.datos.osid.focus();
	return;
}
var d=document.datos;
document.valida.location = "pacientes_con_06.asp?tipo=<%= l_tipo%>&id="+document.datos.id.value + "&dni="+document.datos.dni.value + "&nrohistoriaclinica="+ document.datos.nrohistoriaclinica.value;

//valido();
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.coudes.focus();
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

function Mayuscula(cadena){

	cadena.value = cadena.value.toUpperCase();
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  idciudad "
l_sql = l_sql & " FROM config "
l_sql = l_sql & " where config.empnro = " & Session("empnro")   
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_ciudad     = l_rs("idciudad")
else
   	l_ciudad     = 0
end if
l_rs.Close


select Case l_tipo
	Case "A":
 	    	l_apellido      = ""
	    	l_nombre        = ""
			l_nrohistoriaclinica = "0"
	    	l_dni           = ""
	    	l_domicilio     = ""
			l_telefono      = ""
			l_idobrasocial  = "0"
			l_fecha_ingreso = ""
			l_oblig         = "S"
			l_fechanacimiento = ""
			l_nro_obra_social = ""
			l_sexo = ""
			'l_ciudad  = "0"
			
			l_observaciones = ""
	Case "M":

		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM clientespacientes "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
	    	l_apellido      = l_rs("apellido")
	    	l_nombre        = l_rs("nombre")
			l_nrohistoriaclinica = l_rs("nrohistoriaclinica")
	    	l_dni           = l_rs("dni")
	    	l_domicilio     = l_rs("domicilio")
			l_telefono      = l_rs("telefono")
			l_idobrasocial  = l_rs("idobrasocial")
			l_fecha_ingreso = l_rs("fecha_ingreso") 
			l_oblig         = l_rs("afiliado_obligatorio")
			l_fechanacimiento = l_rs("fechanacimiento")
			l_nro_obra_social = l_rs("nro_obra_social") 
			l_sexo = l_rs("sexo")
			l_ciudad  = l_rs("idciudad")
			if isnull(l_ciudad)  then
				l_ciudad = "0"
			end if
			
			l_observaciones = l_rs("observaciones")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos.apellido.focus();">
<form name="datos" action="pacientes_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="id" value="<%= l_id %>">

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
							<input type="text" name="apellido" size="20" maxlength="20"  onkeydown="Javascript:Mayuscula(this);" value="<%= l_apellido %>">							
						</td>
					    <td align="right"><b>Nombre (*):</b></td>						
						<td>
							<input type="text" name="nombre" size="20" maxlength="20" onkeydown="Javascript:Mayuscula(this);" value="<%= l_nombre %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I. (*):</b></td>
						<td>
							<input type="text" name="dni" size="20" maxlength="20" value="<%= l_dni %>">
						</td>
					    <td align="right"><b>Historia Cl&iacute;nica (*):</b></td>
						<td>
							<input type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">						
						</td>						
					</tr>
					<% If check_genera_histnum(Session("empnro")) then %> 
					<tr>
					    <td align="right"></td>
						<td>							
						</td>
					    <td align="right"><b>Generar Nro.:</b></td>
						<td>														
							<input type=checkbox name="gen_hist_num" size="20" maxlength="20" >							
							<% If l_tipo = "A" then %> <script>document.datos.gen_hist_num.checked=true</script><% End If %> 
							<% If l_nrohistoriaclinica <> "0" and l_nrohistoriaclinica <> "" and IsNumeric(l_nrohistoriaclinica) then %>
								<script>document.datos.gen_hist_num.disabled =true</script>
							<% End If %>							
						</td>						
					</tr>
					<% End If %>
					<tr>
					    <td align="right"><b>Tel&eacute;fono (*):</b></td>
						<td>
							<input type="text" name="telefono" size="20" maxlength="50" value="<%= l_telefono %>">
						</td>
					    <td align="right"><b>Domicilio:</b></td>
						<td >
							<input type="text" name="domicilio" size="30" maxlength="100" value="<%= l_domicilio %>">
						</td>

					</tr>	
					<tr>			
						<td align="right" ><b>Fec. Nacimiento:</b></td>
						<td align="left"  >
						    <input type="text" id="fechanacimiento" name="fechanacimiento" size="10" maxlength="10" value="<%= l_fechanacimiento %>">							
						</td>							
						<td  align="right" nowrap><b>Ciudad: </b></td>
						<td><select name="ciudad" size="1" style="width:150;">
								<option value=0 selected>Seleccione Ciudad</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM ciudades "
								l_sql = l_sql & " where ciudades.empnro = " & Session("empnro")   
								l_sql  = l_sql  & " ORDER BY ciudad "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("ciudad") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.ciudad.value="<%= l_ciudad %>"</script>
						</td>											
					</tr>			
					<tr>
						<td  align="right" nowrap><b>Obra Social (*): </b></td>
						<td><select name="osid" size="1" style="width:200;">
								<option value=0 selected>Seleccione una OS</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM obrassociales "
								l_sql = l_sql & " where obrassociales.empnro = " & Session("empnro")   
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
					    <td align="right"><b>Nro OS:</b></td>
						<td colspan="3">
							<input type="text" name="nro_obra_social" size="20" maxlength="30" value="<%= l_nro_obra_social %>">
						</td>										
					</tr>									
					
					<tr>
					    <td align="right"></td>
						<td>							
						</td>
					    <td align="right"><b>Afiliado Obligatorio:</b></td>
						<td>
							<input  type=checkbox name="afiliado_oblig" size="20" maxlength="20" <% if l_oblig = "S" then %> checked  > <% End If %>
						</td>
					</tr>	
					
					<tr>
					    <td align="right" ><b>Fec. Ingreso:</b></td>
						<td align="left"  >
						    <input type="text" id="fecha_ingreso" name="fecha_ingreso" size="10" maxlength="10" value="<%= l_fecha_ingreso %>">							
						</td>
					    <td align="right"><b>Sexo:</b></td>
						<td ><select name="sexo" size="1" style="width:150;">
								<option value=M selected>Masculino</option>
								<option value="F" >Femenino </option>							
							</select>
							<script>document.datos.sexo.value= "<%= l_sexo %>"</script>						
						</td>						
					</tr>
					
																															
					<tr>
					    <td align="right"><b>Observaciones:</b></td>
						<td colspan="3">
							<input type="text" name="observaciones" size="78" maxlength="200" value="<%= l_observaciones %>">
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
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
