<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Turnos por Empleado.xls" 
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
Dim l_PrecioPractica
Dim l_pagos

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion
Dim l_apeynom

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY calendarios.fechahorainicio DESC "
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
<title>Turnos por Empleado entre Fechas</title>
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
l_idpaciente = request("idpaciente")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo el nombre del Paciente
l_sql = "SELECT  * "
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " WHERE id = " & l_idpaciente
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_apeynom = l_rs("apellido") & ", " & l_rs("nombre")
end if
l_rs.close



l_sql = "SELECT  CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio , CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly ,  recursosreservables.descripcion  " 'calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
'l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono"
'l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.apellido turnoapellido , turnos.nombre turnonombre, turnos.dni turnodni , turnos.domicilio turnodomicilio , turnos.telefono turnotelefono, turnos.comentario turnocomentario"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente, turnos.comentario turnocomentario"
l_sql = l_sql & " FROM turnos "
l_sql = l_sql & " INNER JOIN calendarios ON calendarios.id = turnos.idcalendario "
'l_sql = l_sql & " LEFT JOIN visitas ON visitas.id = practicasrealizadas.idvisita "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = calendarios.idrecursoreservable "
'l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " WHERE turnos.idclientepaciente = " & l_idpaciente
l_sql = l_sql & " AND  calendarios.fechahorainicio  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  calendarios.fechahorainicio <= " & cambiafecha(l_fechahasta,"YMD",true) 

l_sql = l_sql & " and turnos.empnro = " & Session("empnro")   

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
        <td  colspan="6" align="center" ><h3>Turnos desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>
	<tr>
        <td  colspan="6" align="center" ><h3>Paciente:&nbsp;<%= l_apeynom %>&nbsp; </h3></td>	
    </tr>	
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>	
    <tr>
        <th width="100">Fecha - Hora</th>
        <th width="200">Medico</th>			
    </tr>
<% 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Turnos cargados para el filtro ingresado.</td>
</tr>
<%else
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
		
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("DateOnly") %>&nbsp;-&nbsp;<%= l_rs("fechahorainicio") %></td>	
			<td <%'= l_fondo  %> ><%= l_rs("descripcion")%></td>														   
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
