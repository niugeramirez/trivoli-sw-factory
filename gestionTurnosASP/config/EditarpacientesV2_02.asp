
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
  Dim l_hcoblig
  
  l_dnioblig  = request("dnioblig")
  l_hcoblig  = request("hcoblig")

l_ventana = request.querystring("ventana")

'response.write l_tipo
'response.write "l_dnioblig "&l_dnioblig
'response.write "l_hcoblig "&l_hcoblig

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function Validaciones_locales_EditPac_02(){
	//como esta pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local aca, y se invoca desde la ventana llamadora
	var s = document.datos_02_EditPac.osid;

	if (document.datos_02_EditPac.apellido.value == ""){
		alert("Debe ingresar el Apellido del Paciente.");
		document.datos_02_EditPac.apellido.focus();
		return false;
	}

	if (document.datos_02_EditPac.nombre.value == ""){
		alert("Debe ingresar el Nombre del Paciente.");
		document.datos_02_EditPac.nombre.focus();
		return  false;
	}

	 <% If l_dnioblig = "S" then %> 	 
	if (document.datos_02_EditPac.dni.value == "" || document.datos_02_EditPac.dni.value == 0){
		alert("Debe ingresar el DNI del Paciente.");
		document.datos_02_EditPac.dni.focus();
		return false;
	}

	<% End If %>
	if (isNaN(document.datos_02_EditPac.dni.value)) {
		alert("El D.N.I. debe ser numerico.");
		document.datos_02_EditPac.dni.focus();
		return  false;
	}
	if (document.datos_02_EditPac.tel.value == ""){
		alert("Debe ingresar el Telefono del Paciente.");
		document.datos_02_EditPac.tel.focus();
		return false;
	}
	<% if l_hcoblig = "S" then %>
	if ((document.datos_02_EditPac.nrohistoriaclinica.value == "" || document.datos_02_EditPac.nrohistoriaclinica.value == 0) 
			&& document.datos_02_EditPac.gen_hist_num.checked == false){
		alert("Debe ingresar el Nro de Historia Clinica o seleccionar la opcion para generar un nuevo numero.");
		document.datos_02_EditPac.nrohistoriaclinica.focus();
		return  false;
	}

	<% End If %>
	
	if (isNaN(document.datos_02_EditPac.nrohistoriaclinica.value)) {
		alert("El Nro de Historia Clinica debe ser numerico.");
		document.datos_02_EditPac.nrohistoriaclinica.focus();
		return;
	}

	document.datos_02_EditPac.os.value = s.options[s.selectedIndex].text;	
	
	return true;

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
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_EditPac.apellido.focus();">
<form name="datos_02_EditPac" id="datos_02_EditPac" action="Javascript:Submit_Formulario_EditPac();" target="valida">
	<input type="hidden" name="id" value="<%= l_id %>">
	<input type="hidden" name="tipo" value="<%= l_tipo %>">	
	<input type="hidden" name="pacienteid" value="">
	<input type="hidden" name="ventana" value="<%= l_ventana %>">
	<input type="hidden" name="os" value="">

	<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
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
								<script>document.datos_02_EditPac.osid.value="<%= l_idobrasocial %>"</script>
							</td>					
						</tr>
						<tr>
							<td>
							 </br></br>
							</td>
							<td>
							 </br></br>
							</td>						
						</tr>						
						<tr>
							<td align="right"><b> Historia Cl&iacute;nica <% If l_hcoblig = "S" then %> (*)<% End If %>:</b></td>
							<td>
								<input type="text" name="nrohistoriaclinica" size="20" maxlength="20" value="<%= l_nrohistoriaclinica %>">
							</td>					

										
						</tr>	
						<% If check_genera_histnum(Session("empnro")) then %> 						
						<tr>						
							<td align="right"><b>Generar Nro.:</b></td>
							<td>
								<input type=checkbox name="gen_hist_num" size="20" maxlength="20" >
								<% If l_nrohistoriaclinica <> "0" and l_nrohistoriaclinica <> "" and IsNumeric(l_nrohistoriaclinica) then %>
									<script>document.datos_02_EditPac.gen_hist_num.disabled =true</script>
								<% End If %>									
							</td>
										
						</tr>
						<% End If %>						
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	</table>
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
