
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_tipoes
dim l_idtipomovimiento 
dim l_detalle
dim l_idunidadnegocio
dim l_idmediopago
dim l_idcheque
dim l_monto
dim l_idresponsable
dim l_idcompraorigen
dim l_idventaorigen

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
 	    	l_fecha    	   = ""
			l_tipoes    = ""
			'l_fecha_vencimiento     = ""
			l_idtipomovimiento       = "0"
			l_detalle	     = ""
			l_idunidadnegocio    = "0"
			l_idmediopago 		 = "0"
	    	l_idcheque = "0"
	    	l_monto  = "0"
			l_idresponsable = "0"
			l_idcompraorigen = "0"
			l_idventaorigen = "0"

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM cajamovimientos  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		     = l_rs("fecha")
			l_tipoes	     			 = l_rs("tipo")
			l_idtipomovimiento		 = l_rs("idtipomovimiento")
			l_detalle   			 = l_rs("detalle")
			l_idunidadnegocio  		 = l_rs("idunidadnegocio")
			l_idmediopago            = l_rs("idmediopago")
			l_idcheque				 = l_rs("idcheque")
	    	l_monto 				 = l_rs("monto")
	    	l_idresponsable  		 = l_rs("idresponsable")
			l_idcompraorigen         = l_rs("idcompraorigen")
			l_idventaorigen		     = l_rs("idventaorigen")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02.numero.focus();">	
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
								<td align="right"><b>Tipo:</b></td>
								<td colspan="3">
										
									<select name="tipoes" size="1" style="width:250;">		
										<option value= "E" >Entrada</option>
										<option value= "S" >Salida</option>
									</select>
									<script>document.datos_02.tipoes.value= "<%= l_tipoes%>"</script>									
								</td>
				
							</tr>	
							
						    <tr>
								<td align="right"><b>Tipo Movimiento:</b></td>
								<td colspan="3"><select name="idtipomovimiento" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione un Tipo de Movimiento</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM tiposMovimientoCaja "
										l_sql = l_sql & " where tiposMovimientoCaja.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY descripcion "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcion") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idtipomovimiento.value= "<%= l_idtipomovimiento%>"</script>
								</td>					
							</tr>
							
							<tr>
								<td align="right"><b>Detalle:</b></td>
								<td colspan="3">
									<input type="text" name="detalle" size="70" maxlength="200" value="<%= l_detalle %>">							
								</td>
				
							</tr>					
							
						    <tr>
								<td align="right"><b>Unidad de Negocio:</b></td>
								<td colspan="3"><select name="idunidadnegocio" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione una Unidad de Negocio</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM unidadesNegocio "
										l_sql = l_sql & " where unidadesNegocio.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY descripcion "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcion") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idunidadnegocio.value= "<%= l_idunidadnegocio%>"</script>
								</td>					
							</tr>			
							
						    <tr>
								<td align="right"><b>Medio de Pago:</b></td>
								<td colspan="3"><select name="idmediopago" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione un Medio de Pago</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM mediosdepago "
										l_sql = l_sql & " where mediosdepago.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY titulo "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("titulo") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idmediopago.value= "<%= l_idmediopago %>"</script>
								</td>					
							</tr>															
										
						    <tr>
								<td align="right"><b>Cheque:</b></td>
								<td colspan="3"><select name="idcheque" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione un Cheque</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  cheques.id, cheques.numero, bancos.nombre_banco "
										l_sql  = l_sql  & " FROM cheques "
										l_sql  = l_sql  & " LEFT JOIN bancos ON bancos.id = cheques.id_banco "
										l_sql = l_sql & " where cheques.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY bancos.nombre_banco "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre_banco") %> &nbsp;-&nbsp;<%= l_rs("numero") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idcheque.value= "<%= l_idcheque %>"</script>
								</td>					
							</tr>			
							
							<tr>
								<td align="right"><b>Monto:</b></td>
								<td colspan="3">
									<input type="text" name="monto" size="50" maxlength="50" value="<%= l_monto%>">		
									<input type="hidden" name="monto2" value="">						
								</td>
				
							</tr>		
							
						    <tr>
								<td align="right"><b>Responsable:</b></td>
								<td colspan="3"><select name="idresponsable" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione un Responsable</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM responsablesCaja "
										l_sql = l_sql & " where responsablesCaja.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idresponsable.value= "<%= l_idresponsable %>"</script>
								</td>					
							</tr>			
							
						    <tr>
								<td align="right"><b>Compra Origen:</b></td>
								<td colspan="3"><select name="idcompraorigen" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione una Compra</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM compras "
										l_sql  = l_sql  & " LEFT JOIN proveedores ON proveedores.id = compras.idproveedor  "
										l_sql = l_sql & " where compras.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre") %>&nbsp;-&nbsp;<%= l_rs("fecha") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idcompraorigen.value= "<%= l_idcompraorigen %>"</script>
								</td>					
							</tr>																								
																	
																	
							
						    <tr>
								<td align="right"><b>Venta Origen:</b></td>
								<td colspan="3"><select name="idventaorigen" size="1" style="width:250;">
										<option value="0" selected>&nbsp;Seleccione una Venta</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM ventas "
										l_sql  = l_sql  & " LEFT JOIN clientes ON clientes.id = ventas.idcliente  "
										l_sql = l_sql & " where ventas.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre") %>&nbsp;-&nbsp;<%= l_rs("fecha") %> </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idventaorigen.value= "<%= l_idventaorigen %>"</script>
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
