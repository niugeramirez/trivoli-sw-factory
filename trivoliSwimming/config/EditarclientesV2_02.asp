
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
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

dim l_telefono
dim l_celular  
dim l_mail
dim l_direccion
dim l_idciudad

Dim l_ventana

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_id = request.querystring("cabnro")

  Dim l_dnioblig
  Dim l_hcoblig
  
  l_dnioblig  = request("dni")
  l_hcoblig  = request("hcoblig")

l_ventana = request.querystring("ventana")

'response.write l_tipo

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function Validaciones_locales_EditCli_02(){
	//como esta pantalla 02 se usa en varios lugares (a diferencia del esquema general de ABM) ponemos la funcion de validacion local aca, y se invoca desde la ventana llamadora
	
	if (document.datos_02_EditCli.nombre.value == ""){
		alert("Debe ingresar el Nombre del Cliente.");
		document.datos_02_EditCli.nombre.focus();
		return  false;
	}

		
	return true;

}
</script>
<% 
select Case l_tipo
	Case "A":
 	    	
	    	l_nombre        = ""
			l_telefono    = ""
			l_celular     = ""
			l_mail        = ""
			l_direccion   = ""
			l_idciudad    = "0"
	    	
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM clientes "
		'l_sql = l_sql & " INNER JOIN ser_servicio ON ser_servicio.sercod = ser_legajo.legpar1 "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
	    	l_nombre        = l_rs("nombre")
			l_telefono			= l_rs("telefono")
			l_celular			= l_rs("celular")
			l_mail   			= l_rs("mail")
			l_direccion         = l_rs("direccion")
			l_idciudad          = l_rs("idciudad")
			
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_EditCli.nombre.focus();">
<form name="datos_02_EditCli" id="datos_02_EditCli" action="Javascript:Submit_Formulario_EditCli();" target="valida">
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
							
							<td align="right"><b>Nombre (*):</b></td>						
							<td>
								<input type="text" name="nombre" size="20" maxlength="20" value="<%= l_nombre %>">
							</td>						
						</tr>			
							<tr>
								<td align="right"><b>Telefono:</b></td>
								<td colspan="3">
									<input type="text" name="telefono" size="50" maxlength="50" value="<%= l_telefono %>">							
								</td>
				
							</tr>		

							<tr>
								<td align="right"><b>Celular:</b></td>
								<td colspan="3">
									<input type="text" name="celular" size="50" maxlength="50" value="<%= l_celular %>">							
								</td>
				
							</tr>			
							
							<tr>
								<td align="right"><b>Mail:</b></td>
								<td colspan="3">
									<input type="text" name="mail" size="50" maxlength="50" value="<%= l_mail %>">							
								</td>
				
							</tr>																			
			
							<tr>
								<td align="right"><b>Direccion:</b></td>
								<td colspan="3">
									<input type="text" name="direccion" size="50" maxlength="100" value="<%= l_direccion %>">							
								</td>
				
							</tr>	
							
						    <tr>
								<td align="right"><b>Ciudad:</b></td>
								<td colspan="3"><select name="idciudad" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione una Ciudad</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM ciudades "
										' Multiempresa
										' Se agrega este filtro 
										l_sql = l_sql & " where ciudades.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY ciudad "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("ciudad") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_EditCli.idciudad.value= "<%= l_idciudad %>"</script>
								</td>					
							</tr>							
								
												
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
