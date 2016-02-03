
<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_idcliente

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
 	    	l_fecha      = ""
			l_idcliente    = "0"

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM ventas  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha      		= l_rs("fecha")
			l_idcliente         = l_rs("idcliente")
			
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
								<td align="right"><b>Cliente:</b></td>
								<td colspan="3"><select name="idcliente" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Cliente</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM clientes "
										' Multiempresa
										' Se agrega este filtro 
										l_sql = l_sql & " where clientes.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY nombre "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("nombre") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idcliente.value= "<%= l_idcliente %>"</script>
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
