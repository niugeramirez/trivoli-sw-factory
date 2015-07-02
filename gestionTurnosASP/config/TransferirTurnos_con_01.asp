<% Option Explicit %>
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
Dim l_sql2
Dim l_filtro
Dim l_cabnro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant

Dim l_primero

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh  

Dim l_fechadesde
Dim l_fechahasta

Dim l_horadesde
Dim l_horahasta

Dim l_DiasSemana

l_fechadesde = request("fechadesde")
l_fechahasta = request("fechahasta")
'response.write l_fechadesde
'response.write l_fechahasta

l_hd = request("hd")
l_md = request("md")
l_hh = request("hh")
l_mh = request("mh")

l_DiasSemana = request("DiasSemana")

l_horadesde = l_hd & ":" & l_md
l_horahasta = l_hh & ":" & l_mh


l_filtro = request("filtro")
l_cabnro = request("cabnro")
l_orden  = request("orden")


Function diasemana(fecha)


	Select Case weekday(fecha)
		Case 1
			diasemana = "#FFFF80"
		Case 2
			diasemana = "#FF0080"
		Case 3
			diasemana = "#FF1180"
		Case 4
			diasemana = "#FF2280"
		Case 5
			diasemana = "#FF3380"
		Case 6
			diasemana = "#FF4480"
		Case 7
			diasemana = "#FF5580"
	End Select
End Function

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Transferir Turnos</title>
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
        <th>Medico</th>
        <th>Turnos Disponibles</th>
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT   descripcion, recursosreservables.id AS idrecursoreservable, COUNT(*) AS Cantidad "
l_sql = l_sql & " FROM calendarios "
' l_sql = l_sql & " LEFT JOIN turnos ON turnos.idcalendario = calendarios.id "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = calendarios.idrecursoreservable "
'l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = turnos.idos "
'l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = turnos.idpractica "
'l_sql = l_sql & " LEFT JOIN ser_medida       ON ser_legajo.mednro = ser_medida.mednro "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if

l_sql = l_sql & " AND calendarios.id not in (select turnos.idcalendario from turnos)"
l_sql = l_sql & " AND calendarios.estado = 'ACTIVO'"
l_sql = l_sql & " group by descripcion, recursosreservables.id" 
l_sql = l_sql & " " & l_orden


'response.write l_sql
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
		'response.write l_rs("fechahorainicio") '  >= & cambiaformato (Fec,l_horadesde )
	%>
	    <tr  onclick="Javascript:Seleccionar(this,<%= l_rs("idrecursoreservable")%>)">
	        <!--<td width="10%" nowrap><%'= l_rs("buqnro")%></td>		-->
			
	        <td width="10%" nowrap><%= l_rs("descripcion")%></td>		
			  	
			    <td width="10%" nowrap>
				<% 
				l_sql2 = "SELECT id,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS horainicio , fechahorainicio , CONVERT(VARCHAR(10), fechahorainicio, 101) AS Fecha"
				l_sql2 = l_sql2 & " FROM calendarios "
								
				'if l_filtro <> "" then
				  l_sql2 = l_sql2 & " WHERE " & l_filtro & " "
				'end if
				l_sql2 = l_sql2 & " AND calendarios.idrecursoreservable = " & l_rs("idrecursoreservable")
				l_sql2 = l_sql2 & " AND calendarios.id not in (select turnos.idcalendario from turnos)"
				l_sql2 = l_sql2 & " AND calendarios.estado = 'ACTIVO'"
				
				l_sql2 = l_sql2 & " ORDER BY fechahorainicio "
				
				
				'response.write l_sql2&"</br>"
				rsOpen l_rs2, cn, l_sql2, 0
				dim l_a
				%>
				<table border="0" cellpadding="3" cellspacing="3"   >
				<tr> 
				<%
				
				 
				do until l_rs2.eof
					'l_cant = l_cant + 1
					'response.write   instr(l_DiasSemana, weekday (cdate( l_rs2("fechahorainicio"))) )
					
					'and weekday (l_rs2("Fecha")) in l_DiasSemana 
					
					if l_rs2("horainicio") >= l_horadesde  and l_rs2("horainicio") <= l_horahasta and instr(l_DiasSemana, weekday (cdate( l_rs2("fechahorainicio"))) ) <> 0 then
					%>
					<td bgcolor="<%= diasemana(l_rs2("Fecha")) %>" align="center" ><a href="Javascript:parent.abrirVentana('TransferirTurnos_con_02.asp?Tipo=A&ant=<%= l_cabnro %>&nuevo=<%= l_rs2("id")%>' ,'',600,300);"><%= cambiafecha(l_rs2("Fecha"),"","")%><br><%= l_rs2("horainicio")%>&nbsp;<%'= weekday(cdate(l_rs2("Fecha"))) %></a></td>
					<%
					end if
					
					
					l_rs2.MoveNext
				loop		
				l_rs2.Close
				 %>
				 </tr>
				 </table>
				 </td>				   
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
<input type="hidden" name="cabnro" value="<%= l_cabnro %>">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
