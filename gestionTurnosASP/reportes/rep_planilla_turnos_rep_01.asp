<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Planilla de Turnos.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
 %>
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

Dim l_primero
Dim l_fechadesde
Dim l_descripcion

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
<% if request.querystring("excel") = false then  %>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Planilla de Turnos</title>
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
<% 

l_filtro = replace (l_filtro, "*", "%")
l_idrecursoreservable = request("idrecursoreservable")
l_fechadesde = request("qfechadesde")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo la cantidad de turnos simultaneos del Recurso Reservable
l_sql = "SELECT  * "
l_sql = l_sql & " FROM recursosreservables "
l_sql = l_sql & " WHERE id = " & l_idrecursoreservable
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_descripcion = l_rs("descripcion")
	l_cantturnossimult = l_rs("cantturnossimult")
end if
l_rs.close



l_sql = "SELECT  calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono , clientespacientes.nrohistoriaclinica "
l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.apellido turnoapellido , turnos.nombre turnonombre, turnos.dni turnodni , turnos.domicilio turnodomicilio , turnos.telefono turnotelefono, turnos.comentario turnocomentario"
l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.comentario turnocomentario"
l_sql = l_sql & " FROM calendarios "
l_sql = l_sql & " LEFT JOIN turnos ON turnos.idcalendario = calendarios.id "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "
l_sql = l_sql & " WHERE calendarios.idrecursoreservable = " & l_idrecursoreservable
l_sql = l_sql & " AND  CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = " & cambiafecha(l_fechadesde,"YMD",true) 

l_sql = l_sql & " and calendarios.empnro = " & Session("empnro")   

'if l_filtro <> "" then
'  l_sql = l_sql & " WHERE " & l_filtro & " "
'end if
l_sql = l_sql & " " & l_orden

' response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Turno del dia:&nbsp;<%= l_fechadesde %>&nbsp;&nbsp;&nbsp;&nbsp;Dr:&nbsp;<%= l_descripcion %></h3></td>
	
	
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>	
    <tr>
        <th width="100">Hora Desde</th>
        <th width="200">Paciente</th>	
		<th width="200">Nro. Historia Clinica</th>
		<th width="100">Tel&eacute;fono</th>
        <th width="200">Practica</th>	
        <th width="100">Obra Social</th>
        <th width="200">Comentarios</th>		
	
    </tr>
<% 
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
	    <tr  onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>,<%= l_rs("turnoid")%>)">
			
	        <td align="center">
			<% if l_fechahorainicio <> l_rs("fechahorainicio") then 
				l_cantturnos = 1
				response.write l_rs("fechahorainicio") 
			else 
				l_cantturnos = l_cantturnos + 1
				response.write "&nbsp;" 
				
			end if%>
			</td>	
			
			
				
			
			<% if l_rs("estado") = "ANULADO" then ' si esta bloquado: solo desbloquear &acirc; %>
				<td bgcolor="#FFFF80" colspan="5" align="center"  nowrap>Anulado:&nbsp;<%= l_rs("motivo")%></td>	
					
			<% End If %> 
			<% if l_rs("estado") = "ACTIVO" then  ' puede cancelar, transferir  %>
			
			<% if isnull(l_rs("idclientepaciente")) then ' si no esta asignado: asignar, bloquear, borrar %>
			    <td >&nbsp;</td>	
				<td >&nbsp;</td>	
				<td >&nbsp;</td>				
				<td >&nbsp;</td>		
				<td >&nbsp;</td>		
				<td >&nbsp;</td>			
			<% Else  
			
			If clng(l_cantturnos) > clng(l_cantturnossimult) then 
				l_fondo = "bgcolor='#FFDEAD' "
			else 	
				l_fondo = ""
			End If
			
			%>
			
			
			
				<% if l_rs("idclientepaciente") <> -1 then ' si esta asignado a un paciente: cancelar el paciente , transferir %>
			    <td <%= l_fondo  %> ><% If l_fondo <> "" then %>Sobreturno:<% End If %>&nbsp;<%= l_rs("apellido")%>,&nbsp;<%= l_rs("nombre")%></td>	
				<td align="center" <%= l_fondo  %> ><%= l_rs("nrohistoriaclinica")%></td>
				<td <%= l_fondo  %> ><%= l_rs("telefono")%></td>
				<td <%= l_fondo  %> ><%= l_rs("practicanombre")%></td>					
				<td <%= l_fondo  %> ><%= l_rs("osnombre")%></td>		
				<% Else  ' si esta asignado a un paciente Externo (pedir info): cancelar el paciente , transferir  %>
			    <td <%= l_fondo  %> valign="middle"><!--<img src="/turnos/shared/images/mas.png" border="0" alt="Nuevo Paciente">--><% If l_fondo <> "" then %>Sobreturno:<% End If %>&nbsp;<%= l_rs("turnoapellido")%>,&nbsp;<%= l_rs("turnonombre")%></td>	
				<td <%= l_fondo  %> >&nbsp;</td>				
				<td <%= l_fondo  %> ><%= l_rs("turnotelefono")%></td>
				<td <%= l_fondo  %> ><%= l_rs("practicanombre")%></td>					
				<td <%= l_fondo  %> ><%= l_rs("osnombre")%></td>					
				<% End If %>
				
		        <td <%= l_fondo  %> align="center" ><%= l_rs("turnocomentario")%></td>	
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
