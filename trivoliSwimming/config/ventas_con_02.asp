
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_idcliente
Dim l_cliente

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
	buttonImage: "../shared/images/calendar16.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->
</head>

<% 
select Case l_tipo
	Case "A":
 	    	l_fecha      = date()
			l_idcliente    = "0"

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM ventas  "
		l_sql = l_sql & " INNER JOIN clientes ON clientes.id = ventas.idcliente  "
		l_sql  = l_sql  & " WHERE ventas.id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		= l_rs("fecha")
			l_idcliente         = l_rs("idcliente")
			l_cliente           = l_rs("nombre")
			
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
								<td align="right"  ><b>Cliente:</b></td>
								<td>
									<input class="deshabinp" readonly="" type="text" name="cliente" id="cliente" size="20" maxlength="20" value="<%=l_cliente %>">		
									<input type="hidden" name="idcliente2" id="idcliente2" size="10" maxlength="10" value="<%=l_idcliente %>">					
									<a href="Javascript:BuscarCliente();"><img src="../shared/images/Buscar_16.png" border="0" title="Buscar Cliente"></a>	
									<a href="Javascript:Editar_Cliente();"><img src="../shared/images/Modificar_16.png" border="0" title="Editar Cliente"></a>									
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
