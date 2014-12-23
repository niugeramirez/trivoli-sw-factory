<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 
'Archivo: contracts_con_01.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

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

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY fechahorainicio "
end if


'l_ternro  = request("ternro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%> Buques - Buques</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro, turnoid){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	document.datos.idturno.value = turnoid;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Hora Desde</th>
        <th>Paciente</th>	
		<th>Tel&eacute;fono</th>
        <th>Practica</th>	
        <th>Obra Social</th>
        <th>Acciones</th>		
	
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")
l_idrecursoreservable = request("idrecursoreservable")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo la cantidad de turnos simultaneos del Recurso Reservable
l_sql = "SELECT  * "
l_sql = l_sql & " FROM recursosreservables "
l_sql = l_sql & " WHERE id = " & l_idrecursoreservable
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_cantturnossimult = l_rs("cantturnossimult")
end if
l_rs.close



l_sql = "SELECT  calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono"
l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.apellido turnoapellido , turnos.nombre turnonombre, turnos.dni turnodni , turnos.domicilio turnodomicilio , turnos.telefono turnotelefono"
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " LEFT JOIN turnos ON turnos.idcalendario = calendarios.id "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = turnos.idos "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "
'l_sql = l_sql & " LEFT JOIN ser_medida       ON ser_legajo.mednro = ser_medida.mednro "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

' response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Calendarios cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	l_fechahorainicio = ""
	l_cantturnos = 0
	do until l_rs.eof
		l_cant = l_cant + 1
		
	%>
	    <tr  onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>,<%= l_rs("turnoid")%>)">
			
	        <td align="center" width="10%" nowrap>
			<% if l_fechahorainicio <> l_rs("fechahorainicio") then 
				l_cantturnos = 1
				response.write l_rs("fechahorainicio") 
				
				%>
				  
				  <% if l_rs("estado") = "ANULADO" then ' si esta bloquado: solo desbloquear &acirc; %>	
				  <a href="Javascript:parent.abrirVentana('AnularTurno_con_02.asp?Tipo=D&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/candadoabierto.jpg" border="0" alt="Desbloquear Turno"></a>						
				  <% Else  %>
				  	<a href="Javascript:parent.abrirVentana('Asignarpacientes_con_02.asp?Tipo=A&cabnro=' + datos.cabnro.value ,'',600,300);"><img src="/turnos/shared/images/AsignarTurno.png" border="0" alt="Asignar Turno"></a>
				  	<% if l_rs("estado") = "ACTIVO" and isnull(l_rs("idclientepaciente")) then  ' puede cancelar, transferir  %>				  
				  		<a href="Javascript:parent.abrirVentana('AnularTurno_con_02.asp?Tipo=B&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/candado.jpg" border="0" alt="Bloquear Turno"></a>	
				  	<% End If %>	
				  <% End If %>
				  
				  <%
			else 
				l_cantturnos = l_cantturnos + 1
				response.write "&nbsp;" 
				
			end if%>
			</td>	
			
			
				
			
			<% if l_rs("estado") = "ANULADO" then ' si esta bloquado: solo desbloquear &acirc; %>
				<td bgcolor="#FFFF80" colspan="4" align="center" width="10%" nowrap>Anulado:&nbsp;<%= l_rs("motivo")%></td>	
				<td align="center" width="10%" nowrap>
							<!--<a href="Javascript:parent.abrirVentana('AnularTurno_con_02.asp?Tipo=D&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/candadoabierto.jpg" border="0" alt="Desbloquear Turno"></a>			-->
    			</td>					
			<% End If %> 
			<% if l_rs("estado") = "ACTIVO" then  ' puede cancelar, transferir  %>
			
			<% if isnull(l_rs("idclientepaciente")) then ' si no esta asignado: asignar, bloquear, borrar %>
			    <td width="10%" nowrap>&nbsp;</td>	
				<td width="10%" nowrap>&nbsp;</td>					
				<td width="10%" nowrap>&nbsp;</td>		
				<td width="10%" nowrap>&nbsp;</td>				

				
		        <td align="center" width="10%" nowrap>
				                       <!-- <a href="Javascript:parent.abrirVentana('Asignarpacientes_con_02.asp?Tipo=A&cabnro=' + datos.cabnro.value ,'',600,300);"><img src="/turnos/shared/images/AsignarTurno.png" border="0" alt="Asignar Turno"></a> -->
									   <!-- <a href="Javascript:parent.abrirVentana('AnularTurno_con_02.asp?Tipo=B&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/candado.jpg" border="0" alt="Bloquear Turno"></a>			-->
	                                   <!--<a href="Javascript:parent.abrirVentana('EliminarTurnos_con_02.asp?Tipo=A&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/eliminarturno.png" border="0" alt="Eliminar Turno"></a>	-->										   
									   </td>				
			
			<% Else  
			
			If clng(l_cantturnos) > clng(l_cantturnossimult) then 
				l_fondo = "bgcolor='#FFDEAD' "
			else 	
				l_fondo = ""
			End If
			
			%>
			
			
			
				<% if l_rs("idclientepaciente") <> -1 then ' si esta asignado a un paciente: cancelar el paciente , transferir %>
			    <td <%= l_fondo  %> width="10%" nowrap><%= l_rs("apellido")%>,&nbsp;<%= l_rs("nombre")%></td>	
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("telefono")%></td>
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("practicanombre")%></td>					
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("osnombre")%></td>		
				<% Else  ' si esta asignado a un paciente Externo (pedir info): cancelar el paciente , transferir  %>
			    <td <%= l_fondo  %> width="10%" nowrap valign="middle"><img src="/turnos/shared/images/mas.png" border="0" alt="Nuevo Paciente"><%= l_rs("turnoapellido")%>,&nbsp;<%= l_rs("turnonombre")%></td>	
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("turnotelefono")%></td>
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("practicanombre")%></td>					
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("osnombre")%></td>					
				<% End If %>
				
		        <td align="center" width="10%" nowrap>

				                       <a href="Javascript:parent.abrirVentana('CancelarTurnos_con_02.asp?Tipo=A&cabnro=' + datos.cabnro.value + '&turnoid=' + datos.idturno.value,'',600,300);"><img src="/turnos/shared/images/cancelarturno.png" border="0" alt="Cancelar Turno"></a>
	                                   <a href="Javascript:parent.abrirVentana('TransferirTurnos_con_00.asp?Tipo=A&cabnro=<%= l_rs("turnoid")%>' ,'',800,600);"><img src="/turnos/shared/images/transferirturno.png" border="0" alt="Transferir Turno"></a>											   
									   </td>	
			<% End If %>		
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
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="idturno" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
