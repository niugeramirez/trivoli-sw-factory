<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
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

l_filtro = request("filtro")
l_cabnro = request("cabnro")
l_orden  = request("orden")

'if l_orden = "" then
'  l_orden = " ORDER BY fechahorainicio "
'end if


'l_ternro  = request("ternro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
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
        <th>Cant.Turnos</th>		
		<th>Turnos Disponibles</th>
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT   descripcion, COUNT(*) AS Cantidad "
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
l_sql = l_sql & " group by descripcion" 
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
	%>
	    <tr  onclick="Javascript:Seleccionar(this,<%= l_rs("idrecursoreservable")%>)">
	        <!--<td width="10%" nowrap><%'= l_rs("buqnro")%></td>		-->
			
	        <td width="10%" nowrap><%= l_rs("descripcion")%></td>		
			  <td width="10%" nowrap><%= l_rs("Cantidad")%></td>			
			    <td width="10%" nowrap>
				<% 
				l_sql2 = "SELECT id,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS fechahorainicio "
				l_sql2 = l_sql2 & " FROM calendarios "
								
				'if l_filtro <> "" then
				  l_sql2 = l_sql2 & " WHERE " & l_filtro & " "
				'end if
				l_sql2 = l_sql2 & " AND calendarios.id not in (select turnos.idcalendario from turnos)"
				
				l_sql2 = l_sql2 & " ORDER BY fechahorainicio "
				
				
				'response.write l_sql2
				rsOpen l_rs2, cn, l_sql2, 0

				do until l_rs2.eof
					'l_cant = l_cant + 1%>
					<a href="Javascript:parent.abrirVentana('TransferirTurnos_con_02.asp?Tipo=A&ant=<%= l_cabnro %>&nuevo=<%= l_rs2("id")%>' ,'',600,300);"><%= l_rs2("fechahorainicio")%>&nbsp;</a>
					<%
					l_rs2.MoveNext
				loop		
				l_rs2.Close
				 %></td>				   
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
