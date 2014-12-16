<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: cerrar_eval_cap_00.asp
Descripción: Ventana que muestra las Evaluaciones del Evento
Autor : Raul CHinestra
Fecha: 23/01/2004
-->
<% 
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_origen
Dim l_entdes
Dim l_evenro
Dim l_ternro

l_filtro = request("filtro")
l_orden  = request("orden")

'if l_orden = "" then
'  l_orden = " ORDER BY arenro "
'end if

l_evenro = request.querystring("evenro")
l_ternro = request.querystring("ternro")

if isnull(l_ternro) then 
	l_ternro = 0
end if 

'response.write l_ternro
'response.end

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Evaluaciones del Evento - Capacitación - RHPro &reg;</title>
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
        <th>Cod.</th>
        <th>Descripción Mód.</th>		
        <th>Apellido</th>		
		<th>Nombre</th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql =  		 " SELECT cap_falencia.modnro, moddesabr, terape, ternom "
l_sql =  l_sql & " from cap_eventomodulo "
l_sql =  l_sql & " inner join cap_candidato on cap_candidato.evenro = cap_eventomodulo.evenro and "
l_sql =  l_sql & "            cap_candidato.conf = -1 "
l_sql =  l_sql & " inner join cap_falencia on cap_falencia.falorigen <= 6 and "
l_sql =  l_sql & "	          cap_falencia.modnro = cap_Eventomodulo.modnro and "
l_sql =  l_sql & "            cap_falencia.falpendiente = -1 and "
l_sql =  l_sql & "		      cap_falencia.ternro = cap_candidato.ternro "
l_sql =  l_sql & " inner join tercero on tercero.ternro = cap_candidato.ternro "
l_sql =  l_sql & " inner join cap_modulo on cap_modulo.modnro = cap_falencia.modnro "
l_sql =  l_sql & " WHERE cap_eventomodulo.evenro = " & l_evenro 
l_sql =  l_sql & "   AND cap_falencia.ternro = " & l_ternro 

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No Existen Gap Pendientes a Actualizar</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("modnro")%>)">
	        <td width="5%" align="center"><%= l_rs("modnro")%></td>			
	        <td width="32%" align="center"><%= l_rs("moddesabr")%></td>						
    		<td width="32%" align="left"><%= l_rs("terape") %></td>
	        <td width="32%"  align="center" nowrap><%= l_rs("ternom")%></td>
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
