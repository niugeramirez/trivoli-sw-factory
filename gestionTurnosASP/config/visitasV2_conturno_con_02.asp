<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant
dim l_fechahorainicio
dim l_cantturnossimult
dim l_idrecursoreservable
dim l_cantturnos
dim l_fondo
Dim  l_permitir

Dim l_primero

Dim l_calfec

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY fechahorainicio "
end if


'l_ternro  = request("ternro")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="">

	<form name="datos_02_AVCT" id="datos_02_AVCT" action="" method="post" target="valida">
		<table>


			<tr>
				<th>Hora Desde</th>
				<th>Paciente</th>	
				<th>Tel&eacute;fono</th>
				<th>Practica</th>	
				<th>Obra Social</th>
				<th>Asistio </th>		
				<th>No Asistio </th>			
			
			</tr>
		<%
		l_idrecursoreservable = request("idrecursoreservable")
		l_calfec  = request.querystring("fechadesde")


		Set l_rs = Server.CreateObject("ADODB.RecordSet")


		l_sql = "SELECT   calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
		l_sql = l_sql & " , clientespacientes.id clientespacientesid,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono, clientespacientes.nrohistoriaclinica nrohistoriaclinica , clientespacientes.dni dni"
		l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
		l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente"
		l_sql = l_sql & " FROM calendarios "
		l_sql = l_sql & " INNER JOIN turnos ON turnos.idcalendario = calendarios.id "
		l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "
		l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
		l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "
		l_sql = l_sql & " WHERE calendarios.idrecursoreservable =  " & l_idrecursoreservable
		l_sql = l_sql & " AND CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = " & cambiafecha(l_calfec,true,1)  & ""
		l_sql = l_sql & " AND turnos.empnro = " & Session("empnro") 
		l_sql = l_sql & " AND turnos.id NOT IN ( select distinct(idturno) from visitas ) " 
		l_sql = l_sql & " " & l_orden

		 'response.write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		'response.write l_rs.eof
		'response.end

		if l_rs.eof then
			l_primero = 0
		%>
		<tr>
			 <td colspan="7" >No existen Turnos cargados para el filtro ingresado.</td>
		</tr>
		<%else
			l_primero = l_rs("id")
			l_cant = 0
			l_fechahorainicio = ""
			l_cantturnos = 0
			do until l_rs.eof
				l_cant = l_cant + 1
				
			%>
				<tr   onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>,<%= l_rs("turnoid")%>)">
					
					<td align="center" width="10%" nowrap>
					<% if l_fechahorainicio <> l_rs("fechahorainicio") then 
						l_cantturnos = 1
						response.write l_rs("fechahorainicio") 
						else 
						l_cantturnos = l_cantturnos + 1
						response.write "&nbsp;" 
						
						end if%>
					</td>	
					

					
					<% if isnull(l_rs("idclientepaciente")) then %> <!--si no esta asignado: asignar, bloquear, borrar -->
						<td width="10%" nowrap>&nbsp;</td>	
						<td width="10%" nowrap>&nbsp;</td>					
						<td width="10%" nowrap>&nbsp;</td>		
						<td width="10%" nowrap>&nbsp;</td>				

						
						<td align="center" width="10%" nowrap>
											   </td>				
					
					<% Else  
					
					If clng(l_cantturnos) > clng(l_cantturnossimult) then 
						l_fondo = "bgcolor='#FFDEAD' "
					else 	
						l_fondo = ""
					End If
					
					%>
					
					
					
						<% if l_rs("idclientepaciente") <> -1 then %> <!--si esta asignado a un paciente: cancelar el paciente , transferir -->
						<td <%= l_fondo  %> width="10%" nowrap><% If l_rs("nrohistoriaclinica") = "0" or isnull(l_rs("nrohistoriaclinica")) then %>  <img src="/turnos/shared/images/mas.png" border="0" title="Paciente Nuevo"> <% End If %> <%= l_rs("apellido")%>,&nbsp;<%= l_rs("nombre")%></td>	
						<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("telefono")%></td>
						<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("practicanombre")%></td>					
						<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("osnombre")%></td>		
						<% End If 
						
						if isnull(l_rs("dni")) or l_rs("dni") = "" or l_rs("nrohistoriaclinica") = "0" or l_rs("nrohistoriaclinica") = "" or isnull(l_rs("nrohistoriaclinica")) then
						 l_permitir = "N"
						Else
						 l_permitir = "S"
						End If
						%>
									
						<td align="center" width="10%" nowrap><input type=checkbox onclick="Habilitar(this, <%= l_rs("turnoid")%>, '<%= l_permitir %>', 'A')" name="asistio<%= l_rs("turnoid")%>" > </td>    				
						<td align="center" width="10%" nowrap><input type=checkbox onclick="Habilitar2(this, <%= l_rs("turnoid")%>, 'S', 'NA')" name="asistio<%= l_rs("turnoid")%>" value="no"> </td>  
							
													
					<% End If %>		
								   
				</tr>
			<%
				l_fechahorainicio = l_rs("fechahorainicio") 
				l_rs.MoveNext
			loop
		end if

		l_rs.Close
		set l_rs = Nothing
		cn.Close
		set cn = Nothing
		%>

		</table>

		<input type="hidden"  size="400" name="cabnro" value="0">
		<input type="hidden"  size="400" name="cabnro2" value="0">
		<input type="hidden" name="idturno" value="0">
		<input type="hidden" name="orden" value="<%= l_orden %>">
		<input type="hidden" name="filtro" value="<%= l_filtro %>">
	</form>
</body>
</html>
