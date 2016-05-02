
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

Dim l_idcompra

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")
l_idcompra = request.querystring("idcompra")

'response.write l_tipo

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<% 
select Case l_tipo
	Case "A":
			l_idconceptoCompraVenta = "0"
			l_cantidad = ""
			l_precio_unitario = ""
			l_observaciones	 = ""

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM detallecompras  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_idconceptoCompraVenta =  l_rs("idconceptoCompraVenta")
			l_cantidad = l_rs("cantidad")
			l_precio_unitario = l_rs("precio_unitario")
			l_observaciones	 = l_rs("observaciones")		
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_dc.descripcion.focus();">	
	<form name="datos_02_dc" id="datos_02_dc" action = "Javascript:Submit_Formulario_dc();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario_dc();}"  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">
		<input type="Hidden" name="idcompra" value="<%= l_idcompra %>">
		

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
									<script>document.datos_02_dc.idconceptoCompraVenta.value= "<%= l_idconceptoCompraVenta %>"</script>
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
