<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_solicitud_eventos_cap_01.asp
Descripción: Abm de Solicitud de Eventos
Autor : Raul CHinestra
Fecha: 30/03/2004
-->
<% 
on error goto 0

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

'response.write l_ternro

if l_orden = "" then
  l_orden = " ORDER BY solnro "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Solicitud de Eventos - Capacitación - RHPro &reg;</title>
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
        <th>C&oacute;digo</th>
        <th>Descripci&oacute;n</th>		
		<th>Duración</th>
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT solnro,soldesabr,  soldurdias "
l_sql = l_sql & " FROM  cap_solicitud "
l_sql = l_sql & " WHERE cap_solicitud.ternro = " & l_ternro

if l_filtro <> "" then
  l_sql = l_sql & " AND " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="6">No existen Solicitudes de Eventos</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('ag_solicitud_eventos_cap_02.asp?Tipo=M&cabnro=' + datos.cabnro.value+'&ternro=<%= l_ternro %>','',550,200)" onclick="Javascript:Seleccionar(this,<%= l_rs("solnro")%>)">
	        <td width="15%" align="center"><%= l_rs("solnro")%></td>
	        <td width="70%" nowrap><%= l_rs("soldesabr")%></td>
	        <td width="15%"  align="center" nowrap><%= l_rs("soldurdias")%></td>
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
