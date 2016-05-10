
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_idproveedor
dim l_proveedor

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
<!-- Comienzo Datepicker -->
<script>
$(function () {
/*$.datepicker.setDefaults($.datepicker.regional["es"]);
$("#datepicker").datepicker({
firstDay: 1
});*/

		
$( "#fecha" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->
</head>

<% 
select Case l_tipo
	Case "A":
 	    	l_fecha          = date()
			l_idproveedor    = "0"
			l_proveedor      = ""

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  compras.fecha, compras. idproveedor, proveedores.nombre "
		l_sql = l_sql & " FROM compras  "
		l_sql = l_sql & " INNER JOIN proveedores ON proveedores.id = compras.idproveedor  "
		l_sql  = l_sql  & " WHERE compras.id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		= l_rs("fecha")
			l_idproveedor       = l_rs("idproveedor")
			l_proveedor         = l_rs("nombre")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02.fecha.focus();">	
	<form name="datos_02" id="datos_02" action = "Javascript:Submit_Formulario();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario();}"  target="valida">
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
								<td align="right"><b>Fecha:</b></td>
								<td colspan="3">
									<input type="text" id="fecha" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">							
								</td>
				
							</tr>	
												
							

							<tr>
								<td align="right"  ><b>Proveedor:</b></td>
								<td>
									<input class="deshabinp" readonly="" type="text" name="proveedor" id="proveedor" size="20" maxlength="20" value="<%=l_proveedor %>">		
									<input type="hidden" name="idproveedor" id="idproveedor" size="10" maxlength="10" value="<%=l_idproveedor %>">					
									<a href="Javascript:BuscarProveedor();"><img src="../shared/images/Buscar_16.png" border="0" title="Buscar Proveedor"></a>	
									<a href="Javascript:Editar_Proveedor();"><img src="../shared/images/Modificar_16.png" border="0" title="Editar Proveedor"></a>
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
