<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Visitas por Medico entre Fechas.xls" 
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
dim l_idmedio
dim l_cantturnos
dim l_fondo

Dim l_primero
Dim l_fechadesde
Dim l_fechahasta
Dim l_descripcion
Dim l_titulo
Dim l_medico

Dim l_idrecursoreservable
Dim l_idmedicoderivador
Dim l_idpractica

Dim l_preciopractica 
Dim l_PrecioPractica_act
Dim l_monto_pagado 

l_filtro = request("filtro")
l_orden  = request("orden")


'response.end

sub encabezado
 %>
	<tr>
		<td  colspan="11" align="center" ><h3>Medico:&nbsp;<%= l_medico %></h3></td>	
    </tr>	

    <tr>
        <th width="10%">Fecha</th>
        <th width="10%">M&eacute;dico</th>			
        <th width="10%">Paciente</th>	
		<th width="10%">Nro. Historia Clinica</th>			
        <th width="10%">Practica</th>	
		<th width="10%">Medico Derivador</th>
		<th width="10%">Precio Actual</th>
		<th width="10%">Precio Practica</th>
        <th width="10%">Medio Pago</th>
		<th width="10%">OS</th>
		<th width="10%">OS Paciente</th>		
		<th width="10%">Monto Pagado</th>
	
	
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
			<td>&nbsp;</td>			
			<td align="right"><%= l_preciopractica %></td>					
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>		
			<td align="right"><%= l_monto_pagado %></td>								   
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
<title>Pago entre Fechas</title>
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

l_filtro = replace (l_filtro, "*", "%")
l_idmedicoderivador = request("idmedicoderivador")
l_fechadesde = request("qfechadesde")
l_fechahasta = request("qfechahasta")
l_idrecursoreservable = request("idrecursoreservable")

l_idpractica = request("idpractica")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Obtengo el Nombre del Medico
if l_idrecursoreservable = "0" then
	l_medico = "Todos" 
else	
	l_sql = "SELECT  * "
	l_sql = l_sql & " FROM recursosreservables "
	l_sql = l_sql & " WHERE id = " & l_idrecursoreservable
	rsOpen l_rs, cn, l_sql, 0 
	l_medico = ""
	if not l_rs.eof then
		l_medico = l_rs("descripcion")
	end if
	l_rs.close
end if

'response.write  l_idpractica

l_sql = " SELECT visitas.fecha " 
l_sql = l_sql & ",recursosreservables.descripcion medico "
l_sql = l_sql & ",clientespacientes.nombre+' '+clientespacientes.apellido	paciente "
l_sql = l_sql & ",clientespacientes.nrohistoriaclinica "
l_sql = l_sql & ",practicas.descripcion practica "
l_sql = l_sql & ",medicos_derivadores.nombre medico_derivador "
l_sql = l_sql & ",practicasrealizadas.precio preciopractica "
l_sql = l_sql & ",mediosdepago.titulo medio_pago "
l_sql = l_sql & ",obrassociales.descripcion obra_social "
l_sql = l_sql & ",os.descripcion obrasocialpaciente "
l_sql = l_sql & ",sum(pagos.importe) monto_pagado "
l_sql = l_sql & ",visitas.id id_visita "
l_sql = l_sql & ",practicasrealizadas.id id_practicarealizada ,clientespacientes.idobrasocial ,practicasrealizadas.idpractica"
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & "inner join recursosreservables on recursosreservables.id = visitas.idrecursoreservable "
l_sql = l_sql & "inner join clientespacientes on clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & "inner join practicasrealizadas on practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & "inner join practicas on practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & "left join medicos_derivadores on medicos_derivadores.ID = practicasrealizadas.idsolicitadapor "
l_sql = l_sql & "left join pagos on pagos.idpracticarealizada = practicasrealizadas.id "
l_sql = l_sql & "left join mediosdepago on mediosdepago.id = pagos.idmediodepago "
l_sql = l_sql & "left join obrassociales on obrassociales.id = pagos.idobrasocial "
l_sql = l_sql & "left join obrassociales os on os.id = clientespacientes.idobrasocial "
l_sql = l_sql & "where ISNULL(visitas.flag_ausencia ,0) = 0 "
l_sql = l_sql & " and   visitas.fecha  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 
if l_idmedicoderivador <> "0" then
	l_sql = l_sql & " AND medicos_derivadores.id = " & l_idmedicoderivador
end if	
if l_idrecursoreservable <> "0" then
	l_sql = l_sql & " AND recursosreservables.id = " & l_idrecursoreservable
end if	
if l_idpractica <> "0" then
	l_sql = l_sql & " AND  practicas.id = " & l_idpractica
end if	

l_sql = l_sql & " and  visitas.empnro = " & Session("empnro")   
l_sql = l_sql & " group by visitas.fecha "
l_sql = l_sql & ",recursosreservables.descripcion "
l_sql = l_sql & ",clientespacientes.nombre+' '+clientespacientes.apellido "
l_sql = l_sql & ",clientespacientes.nrohistoriaclinica "
l_sql = l_sql & ",practicas.descripcion "
l_sql = l_sql & ",medicos_derivadores.nombre "
l_sql = l_sql & ",practicasrealizadas.precio "
l_sql = l_sql & ",mediosdepago.titulo "
l_sql = l_sql & ",obrassociales.descripcion	"
l_sql = l_sql & ",os.descripcion "
l_sql = l_sql & ",visitas.id "
l_sql = l_sql & ",practicasrealizadas.id ,clientespacientes.idobrasocial ,practicasrealizadas.idpractica "
l_sql = l_sql & "order by id_visita "
l_sql = l_sql & ",id_practicarealizada  "


'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
 %>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <td colspan="12">&nbsp;</td>
    </tr>
	<tr>
        <td  colspan="12" align="center" ><h3>Visitas por Medico desde:&nbsp;<%= l_fechadesde %>&nbsp; al <%= l_fechahasta %>&nbsp;&nbsp;</h3></td>	
    </tr>

<% 	
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="12" >No existen Visitas cargads para el filtro ingresado.</td>
</tr>
<%else
	encabezado
    'l_primero = l_rs("titulo")
	l_preciopractica = 0
	l_monto_pagado = 0
	do until l_rs.eof
		
		'l_cant = l_cant + cdbl(l_rs("importe"))
		
	%>
	    <tr>
			
	        <td align="center"><%= l_rs("fecha") %></td>	
			<td align="left" ><%= l_rs("medico")%></td>	
			<td align="left" ><%= l_rs("paciente")%></td>						
			<td align="center" ><%= l_rs("nrohistoriaclinica")%></td>					
			<td align="left"><%= l_rs("practica")%></td>	
			<td align="left"><%= l_rs("medico_derivador")%></td>	
				<% 				
				l_PrecioPractica_act = PrecioPractica(l_rs("idpractica") , l_rs("idobrasocial") )
				%>				
			<td align="right"><%=  l_PrecioPractica_act%></td>	
			<td align="right"><%= l_rs("preciopractica")%></td>	
			<td align="left"><%= l_rs("medio_pago")%></td>	
			<td align="left"><%= l_rs("obra_social")%></td>
			<td align="left"><%= l_rs("obrasocialpaciente")%></td>
			<td align="right"><%= l_rs("monto_pagado")%></td>								
										   
	    </tr>
	<%
		l_preciopractica = l_preciopractica + cdbl(l_rs("preciopractica"))
		l_monto_pagado = l_monto_pagado + cdbl(l_rs("monto_pagado"))
		
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
