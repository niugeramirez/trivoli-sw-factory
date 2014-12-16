<%
function resumen_objetivossmart
 on error goto 0 
 dim l_evaobjnroant
%>  
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
    <td width="17%">Objetivo</th>
    <td width="17%">Rol</td>
    <td width="17%">Ponderaci&oacute;n</th>
    <td width="17%">% Alcanzado</th>
    <td width="17%">Puntaje</th>
    <td width="17%">Puntaje Manual</th>
    
</tr>	
<%	l_linea = l_linea + 1
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro,evaobjdext, evatipevalua.evatevdesabr, "
   l_sql = l_sql & " evaobjpond, evaobjalcanza, puntaje, puntajemanual, "
   l_sql = l_sql & " evadetevldor.evatevnro, evaoblieva.evaobliorden, evaobjsgto.evasgtotexto "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " AND evadetevldor.evacabnro=" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro "
   	'l_sql = l_sql & " LEFT  JOIN evapuntaje ON evapuntaje.evatevnro = evadetevldor.evatevnro "
   	'l_sql = l_sql & "		AND evapuntaje.evacabnro=" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " INNER JOIN evaoblieva   ON evaoblieva.evatevnro=evadetevldor.evatevnro "
   l_sql = l_sql & "		 AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
   l_sql = l_sql & " INNER  JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & "        AND evasgtotexto is not NULL "
   l_sql = l_sql & " ORDER BY evaobjetivo.evaobjnro "
   rsOpen l_rsl, cn, l_sql, 0
'response.write(l_sql)

 if l_rsl.eof then%>	
	<tr>
	<td valign="top" align="center">--</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	</tr>
<%	l_linea = l_linea + 1
  end if
 l_evaobjnroant=""
 do while not l_rsl.eof %>
 <tr>
	
	<td valign="top" ><b><%if l_evaobjnroant<>l_rsl("evaobjnro") then%><%= l_rsl("evaobjdext")%>&nbsp;<%else%>&nbsp;<%end if%></b></td>	
	<td valign="top" ><b><%= l_rsl("evatevdesabr") %>&nbsp;</b></td>	
 	<td width="15%"><%= l_rsl("evaobjpond") %>&nbsp;</td>
 	<td width="15%"><%= l_rsl("evaobjalcanza") %>&nbsp;</td>
 	<td width="15%"><%= l_rsl("puntaje") %>&nbsp;</td>
 	<td width="15%"><%= l_rsl("puntajemanual") %>&nbsp;</td>
 	</tr>
 	
 	<%if l_evaobjnroant <> l_rsl("evaobjnro") then
 		l_evaobjnroant= l_rsl("evaobjnro")
 	end if
 			
	l_linea = l_linea + 1
	l_rsl.MoveNext
loop%>

			 </table>
			</td>
		 </tr>
  </table>
  <%

l_rsl.Close

end function
%>

