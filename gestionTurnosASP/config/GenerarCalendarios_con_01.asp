<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 
'Archivo: contracts_con_01.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant
Dim l_generar

Dim l_primero

Dim l_fechadesde
Dim l_fechahasta
Dim l_id
Dim l_fecha


l_filtro = request("filtro")
l_orden  = request("orden")

l_generar = request("generar")
l_fechadesde = request("fechadesde")
l_fechahasta = request("fechahasta")
l_id = request("id")


'response.write l_fechadesde
'response.write l_fechahasta
'response.write l_id


Dim l_hd
Dim l_hh
Dim l_hora
Dim aa

Dim l_do 
Dim  l_lu 
Dim l_ma 
Dim l_mi 
Dim l_ju 
Dim l_vi 
Dim l_sa 
Dim l_cm
Dim l_horafin
Dim l_intervaloTurnoMinutos


if l_generar = 1 then

	set l_cm = Server.CreateObject("ADODB.Command")
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT  templatereservasdetalleresumido.id templatereservasdetalleresumido_id, templatereservasdetalleresumido.horainicial, templatereservasdetalleresumido.horafinal , templatereservasdetalleresumido.intervaloTurnoMinutos , * "
	l_sql = l_sql & " FROM recursosreservables "
	l_sql = l_sql & " INNER JOIN templatereservas ON templatereservas.id = recursosreservables.idtemplatereserva "
	l_sql = l_sql & " INNER JOIN templatereservasdetalleresumido ON templatereservasdetalleresumido.idtemplatereserva = templatereservas.id "
	'Eugenio
	l_sql = l_sql & " WHERE recursosreservables.id = "& l_id
	
	l_sql = l_sql & " " & l_orden
	
	'response.write l_sql& "<br>"
	rsOpen l_rs, cn, l_sql, 0 
	do while not l_rs.eof
	
	'response.write "templatereservasdetalleresumido" & l_rs("templatereservasdetalleresumido_id") & "<br>"
	
	l_hd = l_rs("horainicial")
	l_hh = l_rs("horafinal")
	l_do = l_rs("dia1")
	l_lu = l_rs("dia2")
	l_ma = l_rs("dia3")
	l_mi = l_rs("dia4")
	l_ju = l_rs("dia5")
	l_vi = l_rs("dia6")
	l_sa = l_rs("dia7")
	l_intervaloTurnoMinutos = l_rs("intervaloTurnoMinutos")
	
	'response.write "l_intervaloTurnoMinutos" & l_intervaloTurnoMinutos
	'response.end
		
	
	l_fecha = CDate(l_fechadesde)
	Do While DateDiff("d", cdate(l_fecha), CDate(l_fechahasta)) >= 0
	
	'response.write "dia " & cdate(l_fecha) & "<br>" 
	'response.write "l_lu " & l_lu & "<br>"
	'response.write "l_ma " & l_ma & "<br>"
	
	
			 		if (l_lu = "S" and weekday(l_fecha) = 2) or _
					   (l_ma = "S" and weekday(l_fecha) = 3) or _
		 			   (l_mi = "S" and weekday(l_fecha) = 4) or _
					   (l_ju = "S" and weekday(l_fecha) = 5) or _
					   (l_vi = "S" and weekday(l_fecha) = 6) or _
					   (l_sa = "S" and weekday(l_fecha) = 7) or _
					   (l_do = "S" and weekday(l_fecha) = 1) then
					   
					   		'l_caldia = DiadeSemana(l_fecha)
							'aa = DATEDIFF("n", cdate( l_fecha & " " & l_hd ), cdate(l_fecha & " " & l_hh ))
							'response.write "SS" & aa & "<br>"
							
							'response.write "SIIIIII"
							
							l_hora =  l_hd
							l_horafin = DateAdd("n", clng(l_intervaloTurnoMinutos), l_hora)
							
							'response.write "l_hora" & l_hora
							'response.write "l_horafin" & l_horafin
							'response.end
							Do While DATEDIFF("n", cdate( l_horafin ), cdate( l_hh )) >= 0
								
	
								'Verifico que no este repetida el Turno
								l_sql = "SELECT * "
								l_sql = l_sql & " FROM calendarios "
								l_sql = l_sql & " WHERE fechahorainicio=" & cambiaformato (l_fecha,l_hora )
								l_sql = l_sql & " AND fechahorafin=" & cambiaformato (l_fecha,l_horafin )
								l_sql = l_sql & " AND estado='ACTIVO'"
								l_sql = l_sql & " AND idrecursoreservable=" & l_id
								
								rsOpen l_rs2, cn, l_sql, 0
								if not l_rs2.eof then
								    'texto =  "Ya existe otra Obra Social con esa Descripción."
								else
								
									l_sql = "INSERT INTO calendarios "
						            l_sql = l_sql & "(fechahorainicio, fechahorafin, estado, idrecursoreservable ) "
						            l_sql = l_sql & "VALUES (" & cambiaformato (l_fecha,l_hora )  & "," & cambiaformato (l_fecha,l_horafin )  & ",'ACTIVO'," & l_id & ")"
									l_cm.activeconnection = Cn
									l_cm.CommandText = l_sql
						            cmExecute l_cm, l_sql, 0   							
								
								end if 
								l_rs2.close
															
	
								l_hora = DateAdd("n", clng(l_intervaloTurnoMinutos), l_hora)
								l_horafin = DateAdd("n", clng(l_intervaloTurnoMinutos), l_hora)
							Loop
							
							
		       			    'l_sql = "INSERT INTO calendarios "
					        'l_sql = l_sql & "(calfecha, caldia, calhordes, calhorhas, lugnro, evmonro ) "
					        'l_sql = l_sql & "VALUES (" & cambiafecha(l_fecha,"YMD",true)  & ",'" & l_caldia & "','" & l_calhordes1 & l_calhordes2 & "','" & l_calhorhas1 & l_calhorhas2 & "'," & l_lugnro & "," & l_evmonro & ")"
	   					    'l_cm.CommandText = l_sql
					        'cmExecute l_cm, l_sql, 0   

					end if		
	
	  'response.write l_fecha
	  l_fecha = DateAdd("d", 1, l_fecha)
	Loop
	
		l_rs.movenext
	loop
	
end if



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
<title>Generar Calendarios</title>
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
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Hora Desde</th>
		<th>Hora Hasta</th>
        <th>Paciente</th>	
        <th>Practica</th>	
        <th>Obra Social</th>
	
	
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * " ' calendarios.id, estado, motivo,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio, CONVERT(VARCHAR(10), fechahorainicio, 101) AS DateOnly "
l_sql = l_sql & " ,  clientespacientes.apellido, clientespacientes.nombre"
l_sql = l_sql & " ,  obrassociales.descripcion osnombre, practicas.descripcion practicanombre"
'l_sql = l_sql & " ,  turnos.id turnoid, turnos.idclientepaciente, turnos.apellido turnoapellido , turnos.nombre turnonombre, turnos.dni turnodni , turnos.domicilio turnodomicilio"
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
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr  onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
			
	        <td width="10%" nowrap align="center"><%= l_rs("fechahorainicio")%></td>	
			<td width="10%" nowrap align="center"><%= l_rs("fechahorafin")%></td>
			<td width="10%" nowrap>&nbsp;<%= l_rs("apellido")%></td>
			<td width="10%" nowrap>&nbsp;<%= l_rs("practicanombre")%></td>	
			<td width="10%" nowrap>&nbsp;<%= l_rs("osnombre")%></td>				
							   
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
<script>    
	parent.parent.ActPasos(<%= l_primero %>,"","MENU");
    parent.parent.datos.pasonro.value = <%= l_primero %>;
</script>

</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
