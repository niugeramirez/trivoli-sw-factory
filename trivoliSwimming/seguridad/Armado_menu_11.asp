<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden

on error goto 0

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY btnnombre"
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Botones - Ticket</title>
</head>
<script>
var jsSelRow = null;

function Deseleccionar(fila)
{
 fila.className = "MouseOutRow";
}
function Seleccionar(fila,cabnro,pag)
{
 if (jsSelRow != null)
 {
  Deseleccionar(jsSelRow);
 };

 document.datos.cabnro.value = cabnro;
 document.datos.pagina.value = pag;
 fila.className = "SelectedRow";
 jsSelRow		= fila;
}
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th align="left">Nombre</th>
        <th align="left">P&aacute;gina</th>
		<th align="left">Perfiles</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT * "
l_sql = l_sql & "FROM menubtn "
l_sql = l_sql & "WHERE menuraiz = " & request("menuraiz") & " AND menuorder = " & request("menuorder") & " "
if l_filtro <> "" then
  l_sql = l_sql & "AND " & l_filtro & " "
end if
l_sql = l_sql & l_orden
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	do until l_rs.eof
	%>
	    <tr ondblclick="Javascript:parent.abrirVentana('Armado_menu_12.asp?tipo=M&menuraiz=<%= request("menuraiz")%>&menuorder=<%= request("menuorder")%>&cabnro=' + datos.cabnro.value +'&pagina=' + datos.pagina.value,'',500,190)" onclick="Javascript:Seleccionar(this,'<%= l_rs("btnnombre")%>','<%= l_rs("btnpagina")%>')">
	        <td><%= l_rs("btnnombre")%></td>
	        <td><%= l_rs("btnpagina")%></td>
			<td><%= l_rs("btnaccess")%></td>
	    </tr>
	<%
		l_rs.MoveNext
	loop
else%>
		<tr><td colspan="3">No existen Botones Configurados</td></tr>
<%end if
l_rs.Close

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="pagina" value="0">
<input type="Hidden" name="menuraiz" value="<%= request("menuraiz") %>"> 
<input type="Hidden" name="menuorder" value="<%= request("menuorder") %>">
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="Hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
