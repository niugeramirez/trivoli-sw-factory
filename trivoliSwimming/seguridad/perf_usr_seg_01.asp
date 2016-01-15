<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 

'Archivo: perf_usr_seg_01.asp
'Descripción: Abm de Perfiles de usuario
'Autor : Alvaro Bayon
'Fecha: 21/02/2005

Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY perfnom "
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Perfiles de Usuarios - Ticket</title>
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
        <th align="left">Descripci&oacute;n</th>
        <th>Pol&iacute;tica de Cuenta</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT perfnro, perfnom, perftipo, pol_desc "
l_sql = l_sql & "FROM perf_usr LEFT JOIN pol_cuenta ON perf_usr.pol_nro = pol_cuenta.pol_nro "
if l_filtro <> "" then
  l_sql = l_sql & "WHERE " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
	<tr><td colspan="2">No existen Perfiles</td></tr>	
<%else
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('perf_usr_seg_02.asp?Tipo=M&cabnro=' + datos.cabnro.value,'',520,140)" onclick="Javascript:Seleccionar(this,<%= l_rs("perfnro")%>)">
	        <td align="left"><%= l_rs("perfnom")%></td>
	        <td><%= l_rs("pol_desc")%></td>
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
