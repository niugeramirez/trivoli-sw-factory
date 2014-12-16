<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
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

Dim l_primero

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY buqnro "
end if


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
	fila.className = "SelectedRow";
	jsSelRow = fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table>
    <tr>
        <th>Nombre</th>	
        <th>Tipo Operación</th>
        <th>Tipo Buque</th>		
        <th>Agencia</th>		
		<th>Comenzó</th>	
		<th>Terminó</th>			
        <th>Toneladas</th>		
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT    * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " INNER JOIN buq_tipoope   ON buq_tipoope.tipopenro = buq_buque.tipopenro "
l_sql = l_sql & " INNER JOIN buq_tipobuque ON buq_tipobuque.tipbuqnro = buq_buque.tipbuqnro "
l_sql = l_sql & " INNER JOIN buq_agencia   ON buq_agencia.agenro = buq_buque.agenro "
if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then
	l_primero = 0
%>
<tr>
	 <td colspan="7" >No existen Buques cargados para el filtro ingresado.</td>
</tr>
<%else
    l_primero = l_rs("buqnro")
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('buques_con_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',490,300)" onclick="Javascript:Seleccionar(this,<%= l_rs("buqnro")%>)">
	        <!--<td width="10%" nowrap><%'= l_rs("buqnro")%></td>		-->
	        <td width="10%" nowrap><%= l_rs("buqdes")%></td>
	        <td width="10%" nowrap><%= l_rs("tipopedes")%></td>			
	        <td width="10%" nowrap><%= l_rs("tipbuqdes")%></td>						
	        <td width="10%" nowrap><%= l_rs("agedes")%></td>						
	        <td width="10%" nowrap><%= l_rs("buqfecdes")%></td>						
	        <td width="10%" nowrap><%= l_rs("buqfechas")%></td>									
	        <td width="10%" align="right" nowrap><%= l_rs("buqton")%></td>									
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
