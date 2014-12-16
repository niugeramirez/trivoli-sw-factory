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

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY recursosreservables.descripcion "
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
<title>Medicos</title>
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
         <!-- <th>&nbsp;</th>
        <th>Legajo</th>	
        <th>Fec. Ingreso</th>  -->	
        <th>Apellido</th>
        <th>Modelo</th>	
        <th>Cant. Turnos Simultaneos</th>		
        <th>Cant. Sobreturno</th>		
		<!--<th align="left">Domicilio</th>	
		 <th>Derecho Vulnerado</th>  -->			
    </tr>
<%
l_filtro = replace (l_filtro, "*", "%")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    recursosreservables.* , templatereservas.titulo"
l_sql = l_sql & " FROM recursosreservables "
l_sql = l_sql & " LEFT JOIN templatereservas ON templatereservas.id = recursosreservables.idtemplatereserva "
'l_sql = l_sql & " LEFT JOIN ser_medida       ON ser_legajo.mednro = ser_medida.mednro "

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden

'response.write l_sql
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Medicos cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("id")
	l_cant = 0
	do until l_rs.eof
		l_cant = l_cant + 1
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('recursosreservables_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',650,250)" onclick="Javascript:Seleccionar(this,<%= l_rs("id")%>)">
	        <!--<td width="10%" nowrap><%'= l_rs("buqnro")%></td>		-->
			<!-- <td width="2%" nowrap><%'= l_cant %></td>
	        <td width="10%" nowrap><%'= l_rs("legpar1")%>-<%'= l_rs("legpar2")%>/<%'= l_rs("legpar3")%></td>			
	        <td width="10%" nowrap><%'= l_rs("legfecing")%></td>  -->
	        <td width="10%" nowrap><%= l_rs("descripcion")%></td>
			<td width="10%" nowrap><%= l_rs("titulo")%></td>
	        <td width="10%" nowrap><%= l_rs("cantturnossimult")%></td>						
	        <td width="10%" nowrap><%= l_rs("cantsobreturnos")%></td>			
	         <!--<td width="10%" nowrap><%'= l_rs("prodes")%></td>  -->			
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
<!--
<script>    
	parent.parent.ActPasos(<%'= l_primero %>,"","MENU");
    parent.parent.datos.pasonro.value = <%'= l_primero %>;
</script>
-->
</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="0">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
