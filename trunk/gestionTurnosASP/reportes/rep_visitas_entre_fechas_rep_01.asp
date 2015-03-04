<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Visitas entre Fechas.xls" 
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
dim l_idpaciente
dim l_cantturnos
dim l_fondo

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY visitas.fecha "
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
<title>Visitas entre Fechas</title>
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
l_idpaciente = request("idpaciente")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo la cantidad de turnos simultaneos del Recurso Reservable
'l_sql = "SELECT  * "
'l_sql = l_sql & " FROM recursosreservables "
'l_sql = l_sql & " WHERE id = " & l_idrecursoreservable
'rsOpen l_rs, cn, l_sql, 0 
'if not l_rs.eof then
'	l_descripcion = l_rs("descripcion")
'	l_cantturnossimult = l_rs("cantturnossimult")
'end if
'l_rs.close



l_sql = "SELECT  visitas.fecha, practicas.descripcion nombrepractica , recursosreservables.descripcion " 'calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
'l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono"
'l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.apellido turnoapellido , turnos.nombre turnonombre, turnos.dni turnodni , turnos.domicilio turnodomicilio , turnos.telefono turnotelefono, turnos.comentario turnocomentario"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.comentario turnocomentario"
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " LEFT JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
'l_sql = l_sql & " LEFT JOIN visitas ON visitas.id = practicasrealizadas.idvisita "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = visitas.idrecursoreservable "
'l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " WHERE visitas.idpaciente = " & l_idpaciente
l_sql = l_sql & " AND  visitas.fecha  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 

'if l_filtro <> "" then
'  l_sql = l_sql & " WHERE " & l_filtro & " "
'end if
l_sql = l_sql & " " & l_orden

 'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Visitas desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>
	
	
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>	
    <tr>
        <th width="100">Fecha</th>
        <th width="200">Practica</th>	
        <th width="200">Medico</th>	
        
	
	
    </tr>
<% 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Visitas cargadas para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
		
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("fecha") %></td>	
			<td <%'= l_fondo  %> ><%= l_rs("nombrepractica")%>,&nbsp;</td>
			<td <%'= l_fondo  %> ><%= l_rs("descripcion")%></td>					
			
										   
	    </tr>
	<%
		l_rs.MoveNext
	loop
	%>
		 <tr>
			
	        <td align="center">&nbsp;</td>	
			<td align="right">Cantidad</td>					
			<td align="center"><%= l_cant %></td>					
										   
	    </tr>
	<%
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
