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
Dim l_fondoausencia

l_fondovisita = "bgcolor='#FFDEAD' "
l_fondoausencia = "bgcolor='#FFFF80' "

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY calendarios.fechahorainicio "
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
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Paciente</th>
        <th>Nro. Historia Clinica</th>	
		<th>Obra Social</th>	
		<th>Practica</th>
        <th>Precio</th>	
        <th>Precio </br> Pagado</th>
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




l_sql = "SELECT  clientespacientes.id clientespacientesid,  clientespacientes.apellido, clientespacientes.nombre , clientespacientes.nrohistoriaclinica nrohistoriaclinica , isnull(clientespacientes.idobrasocial,0) idobrasocial, obrassociales.descripcion osnombre"
l_sql = l_sql & " , isnull(practicas.id,0) practicaid, practicas.descripcion "
l_sql = l_sql & " ,  visitas.id visitaid , visitas.flag_ausencia "
l_sql = l_sql & " ,  isnull(practicasrealizadas.id,0) practicasrealizadasid , practicasrealizadas.precio "
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN turnos ON turnos.id = visitas.idturno "
l_sql = l_sql & " LEFT JOIN calendarios ON calendarios.id = turnos.idcalendario "


if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
  l_sql = l_sql & " and visitas.empnro = " & Session("empnro")   
else
	l_sql = l_sql & " where visitas.empnro = " & Session("empnro")   
end if
	
l_sql = l_sql & " " & l_orden

'response.write l_sql
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
			
				<% if l_rs("flag_ausencia") = -1 then 
				   		l_fondo = l_fondoausencia
				   Else
				   		l_fondo = l_fondovisita
				   End If 
				%>
	        <tr onclick="Javascript:parent.Seleccionar(this,<%= l_rs("visitaid")%>,document.detalle_01_visit.idvisita);Javascript:parent.Seleccionar(this,0,document.detalle_01_visit.idpracticarealizada);">						
				<td <%= l_fondo %> align="center" width="10%" nowrap><%= l_rs("apellido") %>,&nbsp;<%= l_rs("nombre") %></td>
				<td <%= l_fondo %> align="center" width="10%" nowrap><%= l_rs("nrohistoriaclinica") %></td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>&nbsp;</td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>&nbsp;</td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>&nbsp;</td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>&nbsp;</td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>&nbsp;</td>
				<td <%= l_fondo %> align="center" width="10%" nowrap>
					<% if isnull(l_rs("flag_ausencia")) then %>							
					<a href="Javascript:parent.abrirDialogo('dialogVisit','visitasV2_con_02.asp?Tipo=M&visitaid=<%= l_rs("visitaid") %>' ,950,350);"><img src="../shared/images/Agregar_24.png" border="0" title="Agregar Practica"></a>				                																												
					<% End If %>					
					<a href="Javascript:parent.eliminarRegistroAJAX(<%= l_rs("visitaid") %>,'dialogAlertVisitas','dialogConfirmDeleteVisitas');"><img src="../shared/images/Eliminar_16.png" border="0" title="Eliminar Visita"></a>				
				</td>
			</tr>
			
			<% end if 
			
			if l_rs("practicasrealizadasid") <> 0 then
			%>
			<tr ondblclick="Javascript:parent.abrirDialogo('dialogVisit','visitasV2_con_02.asp?Tipo=M&visitaid=<%= l_rs("visitaid") %>&practicasrealizadasid=<%= l_rs("practicasrealizadasid") %>' ,950,350);" 
				onclick="Javascript:parent.Seleccionar(this,<%= l_rs("practicasrealizadasid")%>,document.detalle_01_visit.idpracticarealizada);parent.Seleccionar(this,0,document.detalle_01_visit.idvisita)">
				<% 
				l_PrecioPractica = PrecioPractica(l_rs("practicaid") ,l_rs("idobrasocial") )
				l_Pagos = Pagos(l_rs("practicasrealizadasid") )
				%>
				<td align="center" width="10%" nowrap>&nbsp;</td>
				<td align="center" width="10%" nowrap>&nbsp;</td>			
				<td align="center" width="10%" nowrap><%= l_rs("osnombre") %></td>		
				<td align="center" width="10%" nowrap><%= l_rs("descripcion") %></td>		
				<td align="center" width="10%" nowrap><%= l_rs("precio") %></td>
				<td align="center" width="10%" nowrap><%= l_Pagos %></td>
				<td align="center" width="10%" nowrap><%= cdbl(l_rs("precio")) - cdbl(l_Pagos) %></td>
				<td align="center" width="10%">													
					<a href="Javascript:parent.abrirDialogo('dialogVisit','visitasV2_con_02.asp?Tipo=M&visitaid=<%= l_rs("visitaid") %>&practicasrealizadasid=<%= l_rs("practicasrealizadasid") %>' ,950,350);"><img src="../shared/images/Modificar_16.png" border="0" title="Editar"></a>
					<a href="Javascript:parent.eliminarRegistroAJAX(<%= l_rs("practicasrealizadasid") %>,'dialogAlertVisitas','dialogConfirmDeleteVisitas');"><img src="../shared/images/Eliminar_16.png" border="0" title="Eliminar Practica"></a>				
					<a href="Javascript:parent.abrirDialogo('dialog_cont_Pagos','pagosV2_con_00.asp?cabnro=<%= l_rs("practicasrealizadasid") %>',620,400);"><img src="../shared/images/US-dollar-icon_16.png" border="0" title="Detalle de Pagos"></a>										
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
<form name="detalle_01_visit" id="detalle_01_visit" method="post">
	<input type="hidden" name="cabnro" value="0">
	<input type="hidden" name="idvisita" value="0">
	<input type="hidden" name="idpracticarealizada" value="0">
	<input type="hidden" name="idturno" value="0">
	<input type="hidden" name="orden" value="<%= l_orden %>">
	<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
