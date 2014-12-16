<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: evento_modulos_cap_01.asp
Descripción: Abm de Modulos asociados al Evento
Autor : Raul CHinestra
Fecha: 05/12/2003
-->
<% 

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_evenro
Dim l_eveorigen
Dim l_eveforeva

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY cap_modulo.modnro "
end if

l_evenro = request.querystring("evenro")


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT eveorigen, eveforeva"
l_sql = l_sql & " FROM cap_evento "
l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	l_eveforeva = l_rs("eveforeva")	
	if isnull(l_rs("eveorigen"))   then 
		l_eveorigen = 0
	else 
		l_eveorigen = l_rs("eveorigen")	
	end if	
	
end if 


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Modulos asociados al Evento - Capacitación - RHPro &reg;</title>
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
 fila.className = "SelectedRow";
 jsSelRow		= fila; 
 
}



</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
	<tr>
		<th  colspan="2" align="center" ><b>Módulos dictados en el Evento</b></th>		
	</tr>				
    <tr>
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>				
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT cap_modulo.modnro,cap_modulo.moddesabr"
l_sql = l_sql & " FROM cap_modulo "
l_sql = l_sql & " INNER JOIN cap_evento ON cap_evento.modnro = cap_modulo.modnro  "
l_sql = l_sql & " WHERE (cap_evento.evenro = " & l_evenro & " AND cap_evento.eveforeva = 1 )"
l_sql = l_sql & " OR (cap_evento.eveorigen = " & l_eveorigen & " AND cap_evento.eveforeva = 2 )"

if l_filtro <> "" then
  l_sql = l_sql & " and " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="6">No existen Módulos asociados al Evento</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr codigo="<%= l_rs("evmonro")%>">
	        <td  align="center" width="15%" align="right"><%= l_rs("modnro")%></td>
	        <td width="85%" nowrap><%= l_rs("moddesabr")%></td>
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
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
