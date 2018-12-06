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
dim l_idos
dim l_cantturnos
dim l_fondo
Dim l_PrecioPractica
Dim l_pagos

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion
Dim l_apeynom

Dim l_nombreOS

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY  visitas.fecha, osnombre, nombremedico "
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


Function PrecioPractica(id_practica , id_obrasocial )
dim l_rs
dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  * "
l_sql = l_sql & " FROM listaprecioscabecera "
l_sql = l_sql & " INNER JOIN listapreciosdetalle ON listapreciosdetalle.idlistaprecioscabecera = listaprecioscabecera.id "
l_sql = l_sql & " WHERE flag_activo = -1 "
l_sql = l_sql & " AND idobrasocial = " & id_obrasocial
l_sql = l_sql & " AND idpractica = " & id_practica
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	PrecioPractica = l_rs("precio")
else
	PrecioPractica = 0
end if
l_rs.close

end Function

Function Pagos(idpracticarealizada )
dim l_rs
dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  * "
l_sql = l_sql & " FROM pagos "
l_sql = l_sql & " WHERE idpracticarealizada = " & idpracticarealizada
rsOpen l_rs, cn, l_sql, 0 
Pagos = 0
do while not l_rs.eof
	Pagos = Pagos + cdbl(l_rs("importe"))
	l_rs.movenext
loop
	
l_rs.close

end Function


l_filtro = replace (l_filtro, "*", "%")
l_idos = request("idos")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")



Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo el Nombre de la OS
if l_idos = "0" then
	l_nombreOS = "Todas"
else	
	l_sql = "SELECT  * "
	l_sql = l_sql & " FROM obrassociales "
	l_sql = l_sql & " WHERE id = " & l_idos
	rsOpen l_rs, cn, l_sql, 0 
	l_nombreOS = ""
	if not l_rs.eof then
		l_nombreOS = l_rs("descripcion")
	end if
	l_rs.close
end if


l_sql = "SELECT  visitas.fecha, practicas.descripcion nombrepractica , recursosreservables.descripcion nombremedico, isnull(practicasrealizadas.id,0) practicasrealizadasid , practicasrealizadas.precio , visitas.flag_ausencia  " 
l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre"
l_sql = l_sql & " ,  obrassociales.descripcion osnombre"

'l_sql = l_sql & " ,  ospago.descripcion ospago "

l_sql = l_sql & " ,  ( select min(ospago.descripcion) from pagos LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial where pagos.idpracticarealizada = practicasrealizadas.id ) ospago "

l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " INNER JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "

'l_sql = l_sql & " LEFT JOIN pagos ON pagos.idpracticarealizada = practicasrealizadas.id "
'l_sql = l_sql & " LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial "

l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = visitas.idrecursoreservable "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " WHERE  visitas.fecha  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & " AND  isnull(visitas.flag_ausencia,0) <> -1" 
l_sql = l_sql & " and visitas.empnro = " & Session("empnro")   

if l_idos <> "0" then
	'l_sql = l_sql &" AND clientespacientes.idobrasocial = "&l_idos
	l_sql = l_sql &" AND ( select min(ospago.id) from pagos LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial where pagos.idpracticarealizada = practicasrealizadas.id ) = "&l_idos

end if
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
        <td  colspan="6" align="center" ><h3>Obra Social:&nbsp;<%= l_nombreOS%>&nbsp; </h3></td>	
    </tr>	
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>	
    <tr>
        <th width="100">Fecha</th>
        <th width="200">OS Paciente</th>	
		<th width="200">Nombre</th>	
		<th width="200">Medico</th>		
        <th width="200">Practica</th>	

		<th width="200">OS Practica</th>		

 		<th width="200">Precio</th>
 		<th width="200">Monto Pagado</th>
 		<th width="200">Saldo</th>  	       
	
	
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
			<td><%= l_rs("osnombre")%></td>	
			<td><%= l_rs("apellido") %>,&nbsp;<%= l_rs("nombre") %></td>	
			<td><%= l_rs("nombremedico")%></td>	
			<% if l_rs("flag_ausencia") = -1 then %>
				<td align="center" >Ausente</td>
				<td align="center" >&nbsp;</td>
				<td align="center" >&nbsp;</td>
				<td align="center" >&nbsp;</td>
			
			<% Else  %>
				<td><%= l_rs("nombrepractica")%>&nbsp;</td>			
				<% 
				l_PrecioPractica = l_rs("precio")
				l_Pagos = Pagos(l_rs("practicasrealizadasid") )
				%>	

				<td  align="right"  ><%= l_rs("ospago") %></td>
		
				<td  align="right"  ><%= l_PrecioPractica %></td>		
				<td  align="right" ><%= l_Pagos	 %></td>	
				<td align="center" ><%= cdbl(l_PrecioPractica) - cdbl(l_Pagos) %></td>			
			<% End If %>			

													   
	    </tr>
	<%
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
