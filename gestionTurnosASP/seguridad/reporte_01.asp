<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim l_rs
Dim l_sql

Dim l_repnro
Dim l_repdesc
dim	l_repagr
	
dim l_filtro
dim l_filtro2
dim l_orden
l_filtro = request("filtro")
l_orden  = request("orden")


l_filtro = request("filtro") 


if len(l_filtro) <> 0 then
	if left(l_filtro,1) <> "'" then
		l_filtro2 = "'" & l_filtro & "'"
	else
		l_filtro2 =  mid(l_filtro,2,len(request("filtro")) - 1)
	end if	
end if	


if l_orden = "" then
		l_orden = " ORDER BY repnro"
end if
%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Reportes - Ticket</title>
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
        <th>Agrupado</th>
    </tr>
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT repnro, "  
l_sql = l_sql & " repdesc,  "
l_sql = l_sql & " repagr   "
l_sql = l_sql & " FROM  reporte"
		
if l_filtro <> "" then
	 l_sql = l_sql & " WHERE " & l_filtro 
end if
	
l_sql = l_sql & " " & l_orden	

rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="2">No hay datos</td>
</tr>
<%else
	do until l_rs.eof
	
	if l_rs("repagr")  then ' es true. INFORMIX CAMBIAR POR -1 ==========
		l_repagr = "agrupado"
	else
		l_repagr = "no agrupado"
	end if	
	
	%>
	<tr onclick="Javascript:Seleccionar(this,<%=l_rs("repnro")%>)">
		<td width="10%"><%=l_rs("repnro")%></td>
		<td width="45%"><%=l_rs("repdesc")%> </td>
		<td width="45%" align=center><%=l_repagr%> </td>
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
