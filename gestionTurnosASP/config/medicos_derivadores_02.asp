<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0


'Datos del formulario
Dim l_id
Dim l_nombre

Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>


<% 
select Case l_tipo
	Case "A":
		l_id = ""
		l_nombre = ""

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM medicos_derivadores "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_nombre = l_rs("nombre")		
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos_med_der.nombre.focus()">	
	<form name="datos_med_der" id="datos_med_der" action = "Javascript:Validar_Formulario();" onkeypress="if (event.keyCode == 13) {event.preventDefault();Validar_Formulario();}"  target="valida">
		<input type="Hidden" name="id" value="<%= l_id %>">
		<input type="Hidden" name="tipo" value="<%= l_tipo %>">


		<table cellspacing="0" cellpadding="0" border="0" width="50%" height="100%">	
		
		<tr>
			<td colspan="2" height="100%">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="50%"></td>
						<td>
						<table cellspacing="0" cellpadding="0" border="0">

							<tr>
								<td align="right"><b>Nombre:</b></td>
								<td> 
									<input type="text" name="nombre" size="30" maxlength="50" value="<%= l_nombre %>">								
								</td>
							</tr>					
						</table>
						</td>
						<td width="50%"></td>
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
