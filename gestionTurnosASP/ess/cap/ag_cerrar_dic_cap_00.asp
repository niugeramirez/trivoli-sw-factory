<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: cerrar_dic_cap_00.asp
Descripción: Ventana que muestra los Objetivos, contenidos y competencias dictados en un Evento
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

l_filtro = request("filtro")
l_orden  = request("orden")

if l_orden = "" then
  l_orden = " ORDER BY arenro "
end if

l_evenro = request.querystring("evenro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Obj, Con, Com dictados en el Evento - Capacitación - RHPro &reg;</title>
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
		<th colspan="4" align="center" class="barra"><b>Objetivos, Contenidos y Competencias Dictados en el Evento</b></th>
	</tr>				
    <tr>
        <th>Origen</th>
		<th>Código</th>		
        <th>Descripci&oacute;n</th>		
		<th> % </th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT origen, evenro, entnro, porcen, objdesabr, temdesabr, evafacdesabr "
l_sql =  l_sql & " from cap_dictado "
l_sql =  l_sql & " left join objetivo on (objetivo.objnro = cap_dictado.entnro and cap_dictado.origen = 1) "
l_sql =  l_sql & " left join tema     on (tema.temnro     = cap_dictado.entnro and cap_dictado.origen = 2) "
l_sql =  l_sql & " left join evafactor    on (evafactor.evafacnro = cap_dictado.entnro and cap_dictado.origen = 3) "
l_sql =  l_sql & " WHERE cap_dictado.evenro = " & l_evenro 

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No se dicto nada en el Evento</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr onclick="Javascript:Seleccionar(this,<%= l_rs("entnro")%>)">
			<% Select Case l_rs("origen")
	    	       Case "1" l_origen = "Objetivo"
				   			l_entdes = l_rs("objdesabr")
		           Case "2" l_origen = "Contenido"
               				l_entdes = l_rs("temdesabr")
		           Case "3" l_origen = "Competencia"
			                l_entdes = l_rs("evafacdesabr")
		      End Select
			 %>
			<td width="30%" align="center"><%= l_origen %></td>
	        <td width="10%" align="center"><%= l_rs("entnro")%></td>			
    		<td width="50%" align="left"><%= l_entdes %></td>
	        <td width="10%"  align="center" nowrap><%= l_rs("porcen")%></td>
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
