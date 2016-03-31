
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id

dim l_fecha_desde 
dim l_fecha_hasta 
dim l_idconceptoCompraVenta 
dim l_cantidadproyectada


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

		
$( "#fecha_desde" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fecha_hasta" ).datepicker({
	showOn: "button",
	buttonImage: "/turnos/shared/images/calendar1.png",
	buttonImageOnly: true
});

});
</script>
<!-- Final Datepicker -->
</head>

<% 
select Case l_tipo
	Case "A":
			l_fecha_desde 				= ""
			l_fecha_hasta  				= ""
			l_idconceptoCompraVenta 	= "0"
			l_cantidadproyectada 		= ""

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM proyeccionventas  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then			
			l_fecha_desde 				= l_rs("fecha_desde")
			l_fecha_hasta  				= l_rs("fecha_hasta")
			l_idconceptoCompraVenta 	= l_rs("idconceptoCompraVenta")
			l_cantidadproyectada 		= l_rs("cantidadproyectada")
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datospv_02.nombre_banco.focus();">	
	<form name="datospv_02" id="datospv_02" action = "Javascript:Submit_Formulario();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario();}"  target="valida">
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
								<td align="right"><b>Fecha Desde:</b></td>
								<td colspan="3">
									<input type="text" name="fecha_desde" id="fecha_desde" size="70" maxlength="200" value="<%= l_fecha_desde %>">							
								</td>
							</tr>
							<tr>
								<td align="right"><b>Fecha Hasta:</b></td>
								<td colspan="3">
									<input type="text" name="fecha_hasta" id="fecha_hasta" size="70" maxlength="200" value="<%= l_fecha_hasta %>">							
								</td>
							</tr>								
							<tr>			 				
								<td align="right"><b>Articulo:</b></td>
								<td colspan="3">									
									<select name="idconceptoCompraVenta" id="idconceptoCompraVenta" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Articulo</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM conceptosCompraVenta "
										' Multiempresa
										' Se agrega este filtro 
										l_sql = l_sql & " where conceptosCompraVenta.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY descripcion "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcion") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>$("#idconceptoCompraVenta").val("<%= l_idconceptoCompraVenta %>")</script>									
								</td>
							</tr>								
							<tr>								
								<td align="right"><b>Cantidad Proyectada:</b></td>
								<td colspan="3">
									<input type="text" name="cantidadproyectada" id="cantidadproyectada" size="70" maxlength="200" value="<%= l_cantidadproyectada %>">							
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
