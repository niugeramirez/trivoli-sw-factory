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

Dim l_PrecioPractica
Dim l_Pagos
Dim l_fondovisita

l_fondovisita = "bgcolor='#FFDEAD' "

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

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	//document.datos.idturno.value = turnoid;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Paciente</th>
        <th>Nro. Historia Clinica</th>	
		<th>Practica</th>
        <th>Precio</th>	
        <th>Precio Pagado</th>
        <th>Saldo</th>	
		<th>Acciones</th>	
	
    </tr>
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

Function Pagos(id_practica , id_obrasocial )
dim l_rs
dim l_sql


Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  * "
l_sql = l_sql & " FROM pagos "
l_sql = l_sql & " WHERE idobrasocial = " & id_obrasocial
l_sql = l_sql & " AND idpracticarealizada = " & id_practica
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	Pagos = l_rs("importe")
else
	Pagos = 0
end if
l_rs.close

	

end Function







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



'l_sql = "SELECT   calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
'l_sql = l_sql & " , clientespacientes.id clientespacientesid,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.telefono, clientespacientes.nrohistoriaclinica nrohistoriaclinica"
'l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
'l_sql = l_sql & " ,  isnull(turnos.id,0) turnoid, turnos.idclientepaciente"
'l_sql = l_sql & " FROM calendarios "
'l_sql = l_sql & " LEFT JOIN turnos ON turnos.idcalendario = calendarios.id "
'l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "
'l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
'l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "
''l_sql = l_sql & " LEFT JOIN ser_medida       ON ser_legajo.mednro = ser_medida.mednro "


l_sql = "SELECT  clientespacientes.id clientespacientesid,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.nrohistoriaclinica nrohistoriaclinica , isnull(clientespacientes.idobrasocial,0) idobrasocial "
l_sql = l_sql & " , isnull(practicas.id,0) practicaid, practicas.descripcion "
l_sql = l_sql & " ,  visitas.id visitaid "
l_sql = l_sql & " ,  isnull(practicasrealizadas.id,0) practicasrealizadasid , practicasrealizadas.precio "
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
'l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
'l_sql = l_sql & " " & l_orden

' response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Visitas cargadas para el filtro ingresado.</td>
</tr>
<%else
    l_primero = 0
	l_cant = 0
	l_fechahorainicio = ""
	l_cantturnos = 0
	do until l_rs.eof
		l_cant = l_cant + 1
		
	%>
	
			<% if l_primero <> l_rs("clientespacientesid") then %>
	        <tr onclick="Javascript:Seleccionar(this,<%= l_rs("visitaid")%>)">
			
			
	        <td <%= l_fondovisita %> align="center" width="10%" nowrap><%= l_rs("apellido") %>,&nbsp;<%= l_rs("nombre") %></td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap><%= l_rs("nrohistoriaclinica") %></td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap>&nbsp;</td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap>&nbsp;</td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap>&nbsp;</td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap>&nbsp;</td>
			<td <%= l_fondovisita %> align="center" width="10%" nowrap>
			<a href="Javascript:parent.abrirVentana('AgregarPractica_con_02.asp?cabnro=<%= l_rs("visitaid") %>' ,'',400,200);"><img src="/turnos/shared/images/colba_24.png" border="0" alt="Agregar Practica"></a>
			<a href="Javascript:parent.abrirVentana('EliminarVisita_con_02.asp?cabnro=<%= l_rs("visitaid") %>' ,'',400,200);"><img src="/turnos/shared/images/eliminarvisita_24.png" border="0" alt="Eliminar Visita"></a>			
			</td>
			</tr>
			
			<% end if 
			
			if l_rs("practicasrealizadasid") <> 0 then
			%>
			<tr>
			<% 
			l_PrecioPractica = PrecioPractica(l_rs("practicaid") ,l_rs("idobrasocial") )
			l_Pagos = Pagos(l_rs("practicasrealizadasid") ,l_rs("idobrasocial") )
			%>
			<td align="center" width="10%" nowrap>&nbsp;</td>
			<td align="center" width="10%" nowrap>&nbsp;</td>
			<td align="center" width="10%" nowrap><%= l_rs("descripcion") %></td>		
			<td align="center" width="10%" nowrap><%= l_rs("precio") %></td>
			<td align="center" width="10%" nowrap><%= l_Pagos %></td>
			<td align="center" width="10%" nowrap><%= cdbl(l_PrecioPractica) - cdbl(l_Pagos) %></td>
			<td align="center" width="10%">
			<a href="Javascript:parent.abrirVentana('AnularTurno_con_02.asp?Tipo=B&cabnro=' + datos.cabnro.value ,'',400,200);"><img src="/turnos/shared/images/money_24.png" border="0" alt="Detalle de Pagos"></a>
			<a href="Javascript:parent.abrirVentana('EliminarPractica_con_02.asp?cabnro=<%= l_rs("practicasrealizadasid") %>' , '',400,200);"><img src="/turnos/shared/images/windows_close_program_24.png" border="0" alt="Eliminar Practicas"></a>
			</td>
			<% End If %>
		    
		</tr>	
<%      
        l_primero = l_rs("clientespacientesid")
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
