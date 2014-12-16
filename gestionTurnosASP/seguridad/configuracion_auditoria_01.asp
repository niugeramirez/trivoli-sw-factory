<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim l_rs
Dim l_sql

Dim l_caudnro
Dim l_cauddes
dim	l_caudact
	
dim l_filtro
dim l_filtro2
dim l_orden
l_filtro = request("filtro")
l_orden  = request("orden")

if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	

if l_orden = "" then
		l_orden = " ORDER BY caudnro"
end if%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Configuraci&oacute;n de Auditor&iacute;s - Ticket</title>
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

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table width="100%" height="100%">
    <tr>
        <th>Conf.Auditor&iacute;a</th>
        <th>Descripci&oacute;n</th>
        <th>Activo</th>
    </tr>
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT caudnro, "  
l_sql = l_sql & " cauddes,  "
l_sql = l_sql & " caudact   "
l_sql = l_sql & " FROM  confaud"
		
if l_filtro <> "" then
	 l_sql = l_sql & " WHERE " & l_filtro 
end if
	
l_sql = l_sql & " " & l_orden	

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="3">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	
	if l_rs("caudact")  then ' es true. INFORMIX CAMBIAR POR -1 ==========
		l_caudact = "activo"
	else
		l_caudact = "inactivo"
	end if	
	
	%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("caudnro")%>)">
		<td width="10%"><%=l_rs("caudnro")%></td>
		<td width="75%"><%=l_rs("cauddes")%> </td>
		<td width="15%" align=center><%=l_caudact%> </td>
	</tr>
	<%l_rs.MoveNext
	loop
end if ' del if l_rs.eof
l_rs.Close
cn.Close	
%>
</table>

<form name="datos" method="post">
<input type="Hidden" name="cabnro" value="0" >
<input type="Hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro2 %>">

</form>

</body>
</html>
