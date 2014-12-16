<%
'  ==============================================================================
'  Para ABN
'  ==============================================================================

function resumen_objetivossmartABN
 dim l_evaobjnroant 
 dim l_evatevnro
 dim l_puntuacionponderada
 dim cols
 cols=4
 
 if l_borrador=-1 then
	cols=cols+1
 end if%>
 
 
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
	<td width="17%">Objetivo</th>
    <td width="17%">Ponderaci&oacute;n</th>
    <td width="17%">% Alcanzado</th>
    <td width="17%">Puntaje</th>
    <%if l_borrador=-1 and l_super<>-1  then%>
    <td width="17%">Observaci&oacute;n</th>
    <%end if%>
</tr>	

<%	l_linea = l_linea + 1
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro,evaobjdext, "
   l_sql = l_sql & " evatipevalua.evatevdesabr,  "
   l_sql = l_sql & " evaobjpond, evaobjalcanza, "
   l_sql = l_sql & " evadetevldor.evatevnro, evaoblieva.evaobliorden, evaobjsgto.evasgtotexto "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & "	    AND evadetevldor.evacabnro=" & l_evacabnro
   if l_borrador=-1 and l_logeadoternro<>"" then
	if l_super<>-1 then
	l_sql = l_sql & "		AND evadetevldor.evaluador = "  & l_logeadoternro
    end if
   end if
   l_sql = l_sql & " LEFT  JOIN evapuntaje ON evapuntaje.evatevnro = evadetevldor.evatevnro "
   l_sql = l_sql & "		AND evapuntaje.evacabnro=" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   if l_borrador<>-1 or l_super=-1 then
   l_sql = l_sql & "		AND evatipevalua.evatevnro = "  & cevaluador
   end if
   l_sql = l_sql & " INNER JOIN evaoblieva   ON evaoblieva.evatevnro=evadetevldor.evatevnro "
   l_sql = l_sql & "		 AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
   l_sql = l_sql & " inner  JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & "        AND evasgtotexto is not NULL "
   l_sql = l_sql & " ORDER BY "
   if l_borrador=-1 then 
   l_sql = l_sql & " evadetevldor.evatevnro, "
   end if
   l_sql = l_sql & " evaobjetivo.evaobjnro "
   rsOpen l_rsl, cn, l_sql, 0
'response.write(l_sql)

 if l_rsl.eof then%>	
	<tr>
	<td valign="top" align="center">--</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<%if l_borrador=-1 then%>
 		<td width="15%">---&nbsp;</td>
 	<%end if%>
	</tr>
<%	l_linea = l_linea + 1
  else
 l_evaobjnroant=""
 l_evatevnro=""
 do while not l_rsl.eof 
	l_evatevnro=l_rsl("evatevnro")
	l_puntuacionponderada = cdbl(l_rsl("evaobjpond")) * cdbl(l_rsl("evaobjalcanza")) / 100
	%>
 <tr>
	<td valign="top" ><b><%if l_evaobjnroant<>l_rsl("evaobjnro") then%><%= l_rsl("evaobjdext")%>&nbsp;<%else%>&nbsp;<%end if%></b></td>	
	<td width="15%"><%= l_rsl("evaobjpond") %>&nbsp;</td>
 	<td width="15%"><%= l_rsl("evaobjalcanza") %>&nbsp;</td>
 	<td width="15%"><%= l_puntuacionponderada %>&nbsp;</td>
 	<%if l_borrador=-1 and l_super<>-1 then%>
 	<td width="15%"><%= l_rsl("evasgtotexto") %>&nbsp;</td>
 	<%end if%>
 </tr>
 	
 	<%if l_evaobjnroant <> l_rsl("evaobjnro") then
 		l_evaobjnroant= l_rsl("evaobjnro")
 	end if
 			
	l_linea = l_linea + 1
	l_rsl.MoveNext
loop
l_rsl.close
set l_rsl=nothing

Set l_rsl = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT puntaje, puntajemanual "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " WHERE evacabnro=" & l_evacabnro
rsOpen l_rsl, cn, l_sql, 0%>

 <tr>
	<td colspan=<%=cols%> valign="top" ><b>Puntuaci&oacute;n de Objetivos (1 a 5) = <%= l_rsl("puntaje") %>
	<% if not isnull(l_rsl("puntaje")) and trim(l_rsl("puntaje"))<>"" and trim(l_rsl("puntajemanual"))<>"" and not isnull(l_rsl("puntajemanual")) then
		' if cdbl(l_rsl("puntaje"))<>cdbl(l_rsl("puntajemanual")) and cint(l_evatevnro)=cevaluador then
		' &nbsp;Modificada a =&nbsp;&nbsp;          = l_rsl("puntajemanual") %>
	<%	' end if
	   end if%>
	</b></td>	
 </tr>
</table>
<%end if
l_rsl.Close
set l_rsl=nothing

end function
%>

