
<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

'Datos del formulario

dim l_id
dim l_fecha
dim l_detalle
dim l_idrecursoreservable
dim l_cantsobreturnos     

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
 	    	l_fecha			     = ""
			l_detalle			 = ""
	    	l_idrecursoreservable = "0"
	    	l_cantsobreturnos  = ""

	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_id = request.querystring("cabnro")
		l_sql = "SELECT  * "
		l_sql = l_sql & " FROM historia_clinica_resumida  "
		l_sql  = l_sql  & " WHERE id = " & l_id
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
 	    	l_fecha     = l_rs("fecha")
			l_detalle   = l_rs("detalle")
	    	l_idrecursoreservable = l_rs("idrecursoreservable")
			
		end if
		l_rs.Close
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="javascript:document.datos_02.fecha.focus();">	
	<form name="datos_02" id="datos_02" action = "Javascript:Submit_Formulario();"   target="valida">
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
									<input type="text" name="fecha" size="10" maxlength="10" value="<%= l_fecha %>">							
								</td>
				
							</tr>	
							
						    <tr>
								<td align="right"><b>Medico:</b></td>
								<td colspan="3"><select name="idrecursoreservable" size="1" style="width:450;">
										<option value="0" selected>&nbsp;Seleccione un Medico</option>
										<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
										l_sql = "SELECT  * "
										l_sql  = l_sql  & " FROM recursosreservables "
										l_sql = l_sql & " where recursosreservables.empnro = " & Session("empnro")   
										
										l_sql  = l_sql  & " ORDER BY descripcion "
										rsOpen l_rs, cn, l_sql, 0
										do until l_rs.eof		%>	
										<option value= <%= l_rs("id") %> > 
										<%= l_rs("descripcion") %>  </option>
										<%	l_rs.Movenext
										loop
										l_rs.Close %>
									</select>
									<script>document.datos_02.idrecursoreservable.value= "<%= l_idrecursoreservable%>"</script>
								</td>					
							</tr>								
							
							<tr>
							    <td align="right"><b>Detalle:</b></td>
								<td>
								    <textarea name="detalle" rows="20" cols="100" ><%= replace(l_detalle,"</br>" , vbCrLf) %></textarea> 
									
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
