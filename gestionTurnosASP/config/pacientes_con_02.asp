
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
dim l_domicilio
dim l_telefono
dim l_idobrasocial
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

if (document.datos.nrohistoriaclinica.value == ""){
	alert("Debe ingresar el Nro de Historia Clinica o ingresar 0.");
	document.datos.nrohistoriaclinica.focus();
	return;
}
/*
if (isNaN(document.datos.nrohistoriaclinica.value)) {
	alert("El Nro de Historia Clinica debe ser numerico.");
	document.datos.nrohistoriaclinica.focus();
	return;
}*/
/*
if (document.datos.domicilio.value == ""){
	alert("Debe ingresar el Domicilio del Paciente.");
	document.datos.domicilio.focus();
	return;
}
*/
if (document.datos.telefono.value == ""){
	alert("Debe ingresar el Telefono del Paciente.");
	document.datos.telefono.focus();
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

</script>
<% 
select Case l_tipo
	Case "A":
 	    	l_apellido      = ""
	    	l_nombre        = ""
			l_nrohistoriaclinica = "0"
	    	l_dni           = ""
	    	l_domicilio     = ""
			l_telefono      = ""
			l_idobrasocial  = "0"
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
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
							<input type="text" name="apellido" size="20" maxlength="20" value="<%= l_apellido %>">							
						</td>
					    <td align="right"><b>Nombre (*):</b></td>						
						<td>
							<input type="text" name="nombre" size="20" maxlength="20" value="<%= l_nombre %>">
						</td>						
					</tr>					
					<tr>
					    <td align="right"><b>D.N.I. (*):</b></td>
						<td>
							<input type="text" name="dni" size="20" maxlength="20" value="<%= l_dni %>">
						</td>
					    <td align="right"><b>Nro. Historia Cl&iacute;nica (*):</b></td>
						<td>
							<input type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
						</td>						
					</tr>
					<tr>
					    <td align="right"><b>Tel&eacute;fono:</b></td>
						<td>
							<input type="text" name="telefono" size="20" maxlength="20" value="<%= l_telefono %>">
						</td>
					    <td align="right"><b>Domicilio:</b></td>
						<td>
							<input type="text" name="domicilio" size="20" maxlength="20" value="<%= l_domicilio %>">
						</td>					
					</tr>			
					<tr>
						<td  align="right" nowrap><b>Obra Social: </b></td>
						<td colspan="3"><select name="osid" size="1" style="width:200;">
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
					-->
					<!-- 
					<tr>
						<td  align="right" nowrap><b>Derecho Vulnerado: </b></td>
						<td colspan="3"><select name="pronro" size="1" style="width:150;">
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
							<script>document.datos.pronro.value= "<%'= l_pronro %>"</script>
						</td>					
					</tr>
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
