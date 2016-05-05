
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_numero
dim l_fecha_emision
dim l_fecha_vencimiento  
dim l_idbanco
dim l_importe
dim l_flag_propio
dim l_flag_emitidopor_cliente
dim l_validacion_bcra
dim l_emisor
dim l_flag_cobrado_pagado

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

		
$( "#fecha_emision" ).datepicker({
	showOn: "button",
	buttonImage: "../shared/images/calendar1.png",
	buttonImageOnly: true
});

$( "#fecha_vencimiento" ).datepicker({
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
 	    	l_numero      	   = ""
			l_fecha_emision    = ""
			l_fecha_vencimiento     = ""
			l_idbanco        = "0"
			l_importe	     = "0"
			l_flag_propio    = "0"
			l_flag_emitidopor_cliente = "0"
			l_emisor 		 = ""
			l_validacion_bcra = "PENDIENTE"
			l_flag_cobrado_pagado = "0"

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		'l_sql = "SELECT  * "

    l_sql = "SELECT    cheques.id ,cheques.numero ,cheques.fecha_emision ,cheques.fecha_vencimiento ,cheques.id_banco ,cheques.importe "
	l_sql = l_sql & " ,cheques.flag_emitidopor_cliente ,cheques.emisor ,cheques.created_by ,cheques.creation_date ,cheques.last_updated_by "
	l_sql = l_sql & " 	,cheques.last_update_date ,cheques.empnro ,ISNULL(cheques.flag_propio,0) as flag_propio , cheques.validacion_bcra "
	l_sql = l_sql & "  ,ISNULL(cheques.flag_cobrado_pagado , 0) as flag_cobrado_pagado "		
		l_sql = l_sql & " FROM cheques  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_numero      		     = l_rs("numero")
			l_fecha_emision		     = l_rs("fecha_emision")
			l_fecha_vencimiento		 = l_rs("fecha_vencimiento")
			l_idbanco   			 = l_rs("id_banco")
			l_importe        		 = l_rs("importe")
			l_flag_propio            = l_rs("flag_propio")
			l_flag_emitidopor_cliente            = l_rs("flag_emitidopor_cliente")	
			l_validacion_bcra				 = l_rs("validacion_bcra")
			l_emisor				 = l_rs("emisor")
			l_flag_cobrado_pagado = l_rs("flag_cobrado_pagado")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02_cheq.numero.focus();">	
	<form name="datos_02_cheq" id="datos_02_cheq" action = "Javascript:Submit_Formulario_cheq();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Submit_Formulario_cheq();}"  target="valida">
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
								<td align="right"><b>Numero:</b></td>
								<td colspan="3">
									<input type="text" name="numero" size="70" maxlength="200" value="<%= l_numero %>">							
								</td>
							</tr>								
							<tr>
								<td align="right"><b>Fecha Emision:</b></td>
								<td colspan="3">
									<input type="text" id="fecha_emision" name="fecha_emision" size="50" maxlength="50" value="<%= l_fecha_emision %>">							
								</td>				
							</tr>									
							<tr>
								<td align="right"><b>Fecha Vencimiento:</b></td>
								<td colspan="3">
									<input type="text" id="fecha_vencimiento" name="fecha_vencimiento" size="50" maxlength="50" value="<%= l_fecha_vencimiento %>">							
								</td>				
							</tr>									
						    <tr>
								<td align="right"><b>Banco:</b></td>
								<td colspan="3"><select name="idbanco" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Banco</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM bancos "
										l_sql = l_sql & " where bancos.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre_banco "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre_banco") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02_cheq.idbanco.value= "<%= l_idbanco%>"</script>
								</td>					
							</tr>		
							
							<tr>
								<td align="right"><b>Importe:</b></td>
								<td colspan="3">
									<input type="text" name="importe" size="50" maxlength="50" value="<%= l_importe%>">		
									<input type="hidden" name="importe2" value="">						
								</td>
				
							</tr>				
						    <tr>
								<td align="right"><b>Emitido por Franquicia:</b></td>
								<td colspan="3"><select name="flag_propio" size="1" style="width:150;">
										<option value="0" selected>NO</option>
										<option value="-1" selected>SI</option>
										
									</select>
									<script>document.datos_02_cheq.flag_propio.value= "<%= l_flag_propio%>"</script>
								</td>					
							</tr>								
						    <tr>
								<td align="right"><b>Emitido por Cliente:</b></td>
								<td colspan="3"><select name="flag_emitidopor_cliente" size="1" style="width:150;">
										<option value="0" selected>NO</option>
										<option value="-1" selected>SI</option>
										
									</select>
									<script>document.datos_02_cheq.flag_emitidopor_cliente.value= "<%= l_flag_emitidopor_cliente%>"</script>
								</td>					
							</tr>		
							
							<tr>
								<td align="right"><b>Emisor:</b></td>
								<td colspan="3">
									<input type="text" name="emisor" size="50" maxlength="50" value="<%= l_emisor %>">							
								</td>
				
							</tr>																								

						    <tr>
								<td align="right"><b>Validacion BCRA:</b></td>
								<td colspan="3"><select name="validacion_bcra" size="1" style="width:150;">
										<option value="PENDIENTE" selected>Pendiente</option>
										<option value="VALIDADO" selected>Validado</option>
										<option value="RECHAZADO" selected>Rechazado</option>
									</select>
									<script>document.datos_02_cheq.validacion_bcra.value= "<%= l_validacion_bcra%>"</script>
								</td>					
							</tr>								

						    <tr>
								<td align="right"><b>Cobrado/Pagado:</b></td>
								<td colspan="3"><select name="flag_cobrado_pagado" size="1" style="width:150;">
										<option value="0" selected>NO</option>
										<option value="-1" selected>SI</option>
										
									</select>
									<script>document.datos_02_cheq.flag_cobrado_pagado.value= "<%= l_flag_cobrado_pagado%>"</script>
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
