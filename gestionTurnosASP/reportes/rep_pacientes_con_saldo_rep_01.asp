<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Pacientes con Saldo.xls" 
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
Dim l_precio_practica
Dim l_monto_pagado
Dim l_monto_deuda
dim l_fechahorainicio
dim l_cantturnossimult
dim l_cantturnos
dim l_fondo

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion
Dim l_titulo


l_filtro = request("filtro")
l_orden  = request("orden")

sub encabezado
 %>

    <tr>
        <th width="100">Fecha</th>
        <th width="200">Paciente</th>	
        <th width="200">Historia Clinica</th>			
        <th width="200">Obra Social</th>	
        <th width="100">Practica</th>
		<th width="100">Medico</th>
		<th width="100">Precio Practica</th>
		<th width="100">Monto Pagado</th>
		<th width="100">Monto Deuda</th>
	
	
    </tr>
<%
end sub	

Sub totales

	%>
		 <tr>
			
	        <td align="center">&nbsp;</td>	
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td align="right">Total</td>					
			<td align="right"><%= l_precio_practica %></td>		
			<td align="right"><%= l_monto_pagado %></td>	
			<td align="right"><%= l_monto_deuda %></td>			
										   
	    </tr>
    	<tr>
        	<td colspan="6">&nbsp;</td>
    	</tr>			
	<%
end sub	

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
<title>Pacientes con Saldo</title>
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

l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT visitas.fecha as fecha_visita " 
l_sql = l_sql & " ,visitas.id as visitasid "
l_sql = l_sql & " ,clientespacientes.id as clientespacientesid " 
l_sql = l_sql & " , clientespacientes.apellido+  ' ' + clientespacientes.nombre as nombre_paciente "
l_sql = l_sql & " , clientespacientes.nrohistoriaclinica  "
l_sql = l_sql & " , obrassociales.descripcion as obra_social "
l_sql = l_sql & " , obrassociales.id as obrassocialesid "
l_sql = l_sql & " , practicas.descripcion as nombrepractica "		
l_sql = l_sql & " , isnull(practicasrealizadas.id,0) as practicasrealizadasid " 	
l_sql = l_sql & " , recursosreservables.id as recursosreservablesid	"
l_sql = l_sql & " , recursosreservables.descripcion as medico "
l_sql = l_sql & " , practicasrealizadas.precio as	precio_practica "
l_sql = l_sql & " , (select ISNULL(sum(pagos.importe),0) from pagos where pagos.idpracticarealizada = practicasrealizadas.id ) as monto_pagado "
l_sql = l_sql & " , practicasrealizadas.precio - (select ISNULL(sum(pagos.importe),0) from pagos where pagos.idpracticarealizada = practicasrealizadas.id ) as monto_deuda "
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " INNER JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " INNER JOIN recursosreservables ON recursosreservables.id = visitas.idrecursoreservable "
l_sql = l_sql & " INNER JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " WHERE ISNULL(visitas.flag_ausencia,0) = 0 "
l_sql = l_sql & " and practicasrealizadas.precio - (select ISNULL(sum(pagos.importe),0) from pagos where pagos.idpracticarealizada = practicasrealizadas.id ) <> 0 "
l_sql = l_sql & " and visitas.fecha >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & " and visitas.empnro = " & Session("empnro")  
l_sql = l_sql & " ORDER BY nombre_paciente,fecha_visita "

 'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Pacientes con Saldo desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>

<% 	
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Pacientes con Saldo cargados para el filtro ingresado.</td>
</tr>
<%else
	encabezado
	l_precio_practica = 0
	l_monto_pagado = 0
	l_monto_deuda = 0
	do until l_rs.eof
				
		l_precio_practica = l_precio_practica + cdbl(l_rs("precio_practica"))
		l_monto_pagado = l_monto_pagado + cdbl(l_rs("monto_pagado"))
		l_monto_deuda = l_monto_deuda + cdbl(l_rs("monto_deuda"))
		
	%>
	    <tr>			
	        <td align="center"><%= l_rs("fecha_visita") %></td>	
			<td ><%= l_rs("nombre_paciente")%></td>
			<td ><%= l_rs("nrohistoriaclinica")%></td>
			<td ><%= l_rs("obra_social")%></td>
			<td ><%= l_rs("nombrepractica")%></td>
			<td ><%= l_rs("medico")%></td>
			<td align="right"><%= l_rs("precio_practica")%></td>
			<td align="right"><%= l_rs("monto_pagado")%></td>
			<td align="right"><%= l_rs("monto_deuda")%></td>
										   
	    </tr>
	<%
		
		l_rs.MoveNext
	loop
	totales

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
