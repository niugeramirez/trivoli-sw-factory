
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
			
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_EditCli.apellido.focus();">
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
