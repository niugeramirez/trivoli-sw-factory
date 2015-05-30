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

Dim l_primero

Dim l_calfec

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
<title>Alta Visitas con Turnos</title>
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
	//document.datos.cabnro.value = document.datos.cabnro.value + "," + cabnro;
	//document.datos.idturno.value = turnoid;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}

function Habilitar(obj, turno){

	if (obj.checked==false) {
	//alert('es falso');
	document.datos.cabnro.value = document.datos.cabnro.value.replace(','+turno, '');
	}
	else {
	//alert('es verdadero ');
	document.datos.cabnro.value = document.datos.cabnro.value + "," + turno ;
	};
	
}

function Validar_Formulario(){

if (Trim(document.datos.cabnro.value) == "0"){
	alert("Debe seleccionar alguna Opcion.");
	document.datos.cabnro.focus();
	return;
}
/*

if (Trim(document.datos.descripcion.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.descripcion.focus();
	return;
}
/*
if (!stringValido(document.datos.agedes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.agedes.focus();
	return;
}
*/
//var d=document.datos;
//document.valida.location = "calendarios_con_06.asp?id=<%= l_id%>&calfec="+document.datos.calfec.value + "&calhordes1="+document.datos.calhordes1.value + "&calhordes2="+document.datos.calhordes2.value + "&calhorhas1="+document.datos.calhorhas1.value + "&calhorhas2="+document.datos.calhorhas2.value + "&intervaloTurnoMinutos="+document.datos.intervaloTurnoMinutos.value ; 


valido();

}

function valido(){
	document.datos.submit();
}


</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">

<form name="datos" action="altavisitaconturno_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<table>


    <tr>
        <th>Hora Desde</th>
        <th>Paciente</th>	
		<th>Tel&eacute;fono</th>
        <th>Practica</th>	
        <th>Obra Social</th>
        <th>Asistio al Turno</th>		
	
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
			    <td <%= l_fondo  %> width="10%" nowrap><% If clng(l_rs("nrohistoriaclinica")) = 0 or isnull(l_rs("nrohistoriaclinica")) then %>  <img src="/turnos/shared/images/mas.png" border="0" alt="Paciente Nuevo"> <% End If %> <%= l_rs("apellido")%>,&nbsp;<%= l_rs("nombre")%></td>	
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("telefono")%></td>
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("practicanombre")%></td>					
				<td <%= l_fondo  %> width="10%" nowrap><%= l_rs("osnombre")%></td>		
				<% End If 
				
				if isnull(l_rs("dni")) or l_rs("dni") = "" or l_rs("nrohistoriaclinica") = "0" or l_rs("nrohistoriaclinica") = "" or isnull(l_rs("nrohistoriaclinica")) then
				%>
					<td align="center" width="10%" nowrap><img src="/turnos/shared/images/cal.gif" border="0" alt="El Paciente seleccionado no tiene DNI o Nro de Historia Clinica cargado. Ir a la opcion Pacientes para completar esta informacion" ></td>
				<% Else  %>
					<td align="center" width="10%" nowrap><input type=checkbox onclick="Habilitar(this, <%= l_rs("turnoid")%>)" name="asistio"> </td>    				
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

<input type="hidden"  size="400" name="cabnro" value="0">
<input type="hidden" name="idturno" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</body>
</html>
