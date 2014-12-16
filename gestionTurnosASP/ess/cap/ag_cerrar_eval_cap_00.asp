<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo: cerrar_eval_cap_00.asp
Descripci�n: Ventana que muestra las Evaluaciones del Evento
Autor : Raul CHinestra
Fecha: 23/01/2004
-->
<% 
'on error goto 0
Dim l_rs
Dim l_sql
Dim l_filtro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_origen
Dim l_entdes
Dim l_evenro
Dim l_eveforeva
Dim l_eveorigen

l_filtro = request("filtro")
l_orden  = request("orden")



Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_evenro = request.querystring("evenro")
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
<title>Evaluaciones del Evento - Capacitaci�n - RHPro &reg;</title>
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
        <th>Tipo</th>
		<th>Formador</th>		
        <th>Participante</th>		
		<th> Resultado </th>		
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql =         " SELECT  tercero.terape, tercero.ternom, evatipo, formador.terape, formador.ternom, resdesabr, evaporcen "
l_sql = l_sql & " from cap_evaluacion "
l_sql = l_sql & " Inner join tercero on tercero.ternro = cap_evaluacion.evaparticipante "
l_sql = l_sql & " Inner join tercero as formador on formador.ternro = cap_evaluacion.evaformador "
l_sql = l_sql & " Inner join resultado on resultado.resnro = cap_evaluacion.evaresnro##1 "

l_sql = l_sql & " WHERE (cap_evaluacion.evenro = " & l_evenro & " AND cap_evaluacion.evento_origen = 1 )"
l_sql = l_sql & " OR (cap_evaluacion.evenro = " & l_eveorigen & " AND cap_evaluacion.evento_origen = 2 )"

if l_filtro <> "" then
  l_sql = l_sql & " WHERE " & l_filtro & " "
end if
l_sql = l_sql & " " & l_orden
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
<tr>
	 <td colspan="4">No Hay Evaluaciones</td>
</tr>
<%else
	do until l_rs.eof
	%>
	    <tr >
			<% Select Case l_rs("evatipo")
	    	       Case "3" l_origen = 3
				   			l_entdes = "Satisfacci�n"
		           Case "4" l_origen = 4
               				l_entdes = "Conoc."
		           Case "5" l_origen = 5
			                l_entdes = "Aplic."
		      End Select
			 %>
			<td width="5%" align="center"><%= l_entdes %></td>
	        <td width="40%" align="left"><%=  l_rs(4)%>,<%= l_rs(5)%> </td>			
	        <td width="40%" align="left"><%=  l_rs(1)%>,<%= l_rs(2)%> </td>			
	        <td width="10%"  align="center" nowrap><%= l_rs("resdesabr") %></td>
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
