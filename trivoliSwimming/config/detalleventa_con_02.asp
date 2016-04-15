
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_idconceptoCompraVenta
dim l_cantidad
dim l_precio_unitario
dim l_observaciones
dim l_idestadoInstalacion 
dim l_fechaProgramadaInstalacion

Dim l_idventa

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_idventa = request.querystring("idventa")

'response.write l_tipo

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Comienzo Datepicker -->
<script>
$(function () {
		
$( "#fechaProgramadaInstalacion" ).datepicker({
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
			l_idconceptoCompraVenta = "0"
			l_cantidad = ""
			l_precio_unitario = ""
			l_observaciones	 = ""
			l_idestadoInstalacion = ""
			l_fechaProgramadaInstalacion = ""
		
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  * "
			l_sql = l_sql & " FROM estadoInstalacion  "
			l_sql = l_sql & " where estadoInstalacion.empnro = " & Session("empnro")   										
			l_sql  = l_sql  & " ORDER BY estadoInstalacion.orden "
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
				l_idestadoInstalacion = l_rs("id")	
			else
				l_idestadoInstalacion = "0"
			end if
			l_rs.Close
			
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM detalleVentas  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idconceptoCompraVenta =  l_rs("idconceptoCompraVenta") 
			l_cantidad = l_rs("cantidad")
			l_precio_unitario = l_rs("precio_unitario")
			l_observaciones	 = l_rs("observaciones")				
			if isnull(l_rs("idestadoInstalacion"))  then 			
				l_idestadoInstalacion = "0" 
			else 			
				l_idestadoInstalacion = l_rs("idestadoInstalacion")
			end if
			l_fechaProgramadaInstalacion = l_rs("fechaProgramadaInstalacion")			
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_dv.descripcion.focus();">	
	<form name="datos_02_dv" id="datos_02_dv" action = "Javascript:Submit_Formulario_dv();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario_dv();}"  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">
		<input type="Hidden" name="idventa" value="<%= l_idventa %>">
		

		<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
		<tr>
			<td colspan="2" height="100%">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<table cellspacing="0" cellpadding="0" border="0">
				
							
						    <tr>
								<td align="right"><b>Concepto:</b></td>
								<td colspan="3"><select name="idconceptoCompraVenta" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Concepto</option>
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
									<script>document.datos_02_dv.idconceptoCompraVenta.value= "<%= l_idconceptoCompraVenta %>"</script>
								</td>					
							</tr>

							

							<tr>
								<td align="right"><b>Cantidad:</b></td>
								<td colspan="3">
									<input type="text" name="cantidad" size="50" maxlength="50" value="<%= l_cantidad %>">
									<input type="hidden" name="cantidad2" value="">								
								</td>
				
							</tr>		

							<tr>
								<td align="right"><b>Precio Unitario:</b></td>
								<td colspan="3">
									<input type="text" name="precio_unitario" size="50" maxlength="50" value="<%= l_precio_unitario%>">		
									<input type="hidden" name="precio_unitario2" value="">						
								</td>
				
							</tr>			
							<tr>
								<td align="right"><b>Estado Inslacion:</b></td>
								<td colspan="3">									
									<select name="estadoInstalacion" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Estado</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM estadoInstalacion "
										' Multiempresa
										' Se agrega este filtro 
										l_sql = l_sql & " where estadoInstalacion.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY orden "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcionEstadoInsta") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_dv.estadoInstalacion.value= "<%= l_idestadoInstalacion %>"</script>
								</td>
				
							</tr>
							<tr>
								<td align="right"><b>Fecha Programada Instalacion:</b></td>
								<td colspan="3">
									<input type="text" id="fechaProgramadaInstalacion"  name="fechaProgramadaInstalacion" size="50" maxlength="50" value="<%= l_fechaProgramadaInstalacion %>">							
								</td>
				
							</tr>							
							<tr>
								<td align="right"><b>Observaciones:</b></td>
								<td colspan="3">
									<input type="text" name="observaciones" size="50" maxlength="50" value="<%= l_observaciones %>">							
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
