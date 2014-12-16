<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_evaluar_eventos_cap_01.asp
Descripción: Abm de Evaluaciones de Eventos
Autor : Raul CHinestra (listo)
Fecha: 21/06/2007
-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_ternro

l_filtro = request("filtro")
l_orden  = request("orden")
l_ternro  = l_ess_ternro

if l_orden = "" then
  l_orden = " ORDER BY evecodext "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaciones de Eventos - Capacitación - RHPro &reg;</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.ttesnro.value = fila.codigo;

 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>		
		<th>Tipo de Test</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT cap_evento.evenro, evecodext, evedesabr, pos_tipotest.ttesnro, ttesdesabr  "
l_sql = l_sql & " FROM  cap_candidato "
l_sql = l_sql & " INNER JOIN cap_evento ON cap_candidato.evenro = cap_evento.evenro "
l_sql = l_sql & " INNER JOIN cap_eventotipotest ON cap_eventotipotest.evenro = cap_evento.evenro "
l_sql = l_sql & " INNER JOIN pos_tipotest ON pos_tipotest.ttesnro = cap_eventotipotest.ttesnro "
l_sql = l_sql & " WHERE cap_candidato.conf = -1 "
l_sql = l_sql & "   AND cap_candidato.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No existen Evaluaciones pendientes a Eventos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr codigo="<%= l_rs("ttesnro")%>" ondblclick="Javascript:parent.abrirVentana('ag_evaluar_eventos_cap_02.asp?ttesnro='+document.datos.ttesnro.value+'&ternro='+ document.datos.ternro.value+'&evenro='+ document.datos.cabnro.value,'',780,580)" onclick="Javascript:Seleccionar(this,<%= l_rs("evenro")%>)">
	        <td width="20%" align="center"><%= l_rs("evecodext")%></td>
	        <td width="40%" nowrap><%= l_rs("evedesabr")%></td>
	        <td width="40%"  align="center" nowrap><%= l_rs("ttesdesabr")%></td>
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
<input type="hidden" name="ttesnro" value="0">
<input type="hidden" name="ternro" value="<%= l_ternro %>">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
