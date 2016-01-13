<% function form_objetivossmart %>  
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
    <td width="17%">Evaluador</td>
    <td width="17%">Descripci&oacute;n</th>
    <td width="17%">Ponderaci&oacute;n</th>
    <td width="17%">% Alcanzado</th>
    <td width="17%">Observaci&oacute;n</th>
    <td width="17%">Puntaje</th>
    <td width="17%">Puntaje Manual</th>
    
</tr>	
<%
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro,evaobjdext, "
   l_sql = l_sql & " evatipevalua.evatevdesabr,  "
   l_sql = l_sql & " evaobjpond,  "
   l_sql = l_sql & " evadetevldor.evatevnro, evaoblieva.evaobliorden  "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " INNER JOIN evacab ON evadetevldor.evacabnro = evacab.evacabnro "
   'l_sql = l_sql & " LEFT  JOIN evapuntaje ON evapuntaje.evatevnro = evadetevldor.evatevnro "
   'l_sql = l_sql & "		AND evapuntaje.evacabnro=" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " INNER JOIN evaoblieva   ON evaoblieva.evatevnro=evadetevldor.evatevnro "
   l_sql = l_sql & "		 AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
   l_sql = l_sql & " INNER JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " AND evacab.evaevenro=" & l_evaevenro
   l_sql = l_sql & " AND evacab.empleado=" & l_ternro
   l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
   rsOpen l_rsl, cn, l_sql, 0

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
<%end if

 do while not l_rsl.eof %>
 <tr>
	<td valign="top" ><b><%= l_rsl("evatevdesabr") %>&nbsp;</b></td>	
 	<td valign="top" ><b><%= l_rsl("evaobjdext") %>&nbsp;</b></td>	
 	<td width="15%"><%= l_rsl("evaobjpond") %>&nbsp;</td>
 	<td width="15%">&nbsp;</td>
 	<td width="15%">&nbsp;</td>
 	<td width="15%">&nbsp;</td>
 	<td width="15%">&nbsp;</td>
 </tr>
<%l_rsl.MoveNext
loop%>

			 </table>
			</td>
		 </tr>
  </table>
  <%

l_rsl.Close
set l_rsl=nothing

end function
%>

