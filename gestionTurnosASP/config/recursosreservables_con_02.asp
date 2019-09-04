
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_descripcion
dim l_idtemplatereserva
dim l_cantturnossimult  
dim l_cantsobreturnos     
dim l_nromatricula

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
 	    	l_descripcion       = ""
			l_idtemplatereserva = "0"
	    	l_cantturnossimult  = ""
	    	l_cantsobreturnos   = ""
			l_nromatricula      = "" 
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM recursosreservables  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_descripcion       = l_rs("descripcion")
			l_idtemplatereserva = l_rs("idtemplatereserva")
	    	l_cantturnossimult  = l_rs("cantturnossimult")
	    	l_cantsobreturnos   = l_rs("cantsobreturnos")
			l_nromatricula      = l_rs("nro_matricula") 
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02.descripcion.focus();">	
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
								<td align="right"><b>Apellido:</b></td>
								<td colspan="3">
									<input type="text" name="descripcion" size="37" maxlength="37" value="<%= l_descripcion %>">							
								</td>
							</tr>	
							<tr>
								<td align="right"><b>Modelo:</b></td>
								<td colspan="3"><select name="idtemplatereserva" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Modelo</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM templatereservas "
										' Multiempresa
										' Se agrega este filtro 
										l_sql = l_sql & " where templatereservas.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY titulo "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										(<%= l_rs("descripcion") %> ) <%= l_rs("titulo") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idtemplatereserva.value= "<%= l_idtemplatereserva %>"</script>
								</td>					
							</tr>											
							<tr>
								<td align="right"><b>Cant. Turnos Simultaneos:</b></td>
								<td>
									<input type="text" name="cantturnossimult" size="20" maxlength="20" value="<%= l_cantturnossimult %>">
								</td>
							</tr>
							<tr>
								<td align="right"><b>Matricula:</b></td>
								<td>
									<input type="text" name="nromatricula" size="20" maxlength="20" value="<%= l_nromatricula %>">
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
