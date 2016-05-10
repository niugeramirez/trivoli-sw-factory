
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_nombre
dim l_telefono
dim l_celular  
dim l_mail

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

'response.write l_tipo

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<% 
select Case l_tipo
	Case "A":
 	    	l_nombre      = ""
			l_telefono    = ""
			l_celular     = ""
			l_mail        = ""


	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM proveedores  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_nombre      		= l_rs("nombre")
			l_telefono			= l_rs("telefono")
			l_celular			= l_rs("celular")
			l_mail   			= l_rs("mail")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_prov.nombre.focus();">	
	<form name="datos_02_prov" id="datos_02_prov" action = "Javascript:Submit_Formulario_prov();" onkeypress=""  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">

		<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
		<tr>
			<td colspan="2" height="100%">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<table cellspacing="0" cellpadding="0" border="0">
				
															
							<tr>
								<td align="right"><b>Nombre:</b></td>
								<td colspan="3">
									<input type="text" name="nombre" name="nombreproveedor" size="70" maxlength="200" value="<%= l_nombre %>">							
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
