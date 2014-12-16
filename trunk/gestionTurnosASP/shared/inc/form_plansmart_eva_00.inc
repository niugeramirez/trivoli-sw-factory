<% 
function form_plansmart
 Dim l_evatevdesabr	
 Dim l_evatevnro
 Dim l_rs_pa
 
   Set l_rs_pa = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro,evaobjdext, "
   l_sql = l_sql & " evatipevalua.evatevdesabr,  "
   l_sql = l_sql & " evadetevldor.evatevnro, evaoblieva.evaobliorden  "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " INNER JOIN evacab ON evadetevldor.evacabnro = evacab.evacabnro "
   l_sql = l_sql & " LEFT  JOIN evapuntaje ON evapuntaje.evatevnro = evadetevldor.evatevnro "
   l_sql = l_sql & "		AND evapuntaje.evacabnro=" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " INNER JOIN evaoblieva   ON evaoblieva.evatevnro=evadetevldor.evatevnro "
   l_sql = l_sql & "		 AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
   l_sql = l_sql & " INNER JOIN evaobjsgto ON evaobjsgto.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & "        AND evaobjsgto.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " AND evacab.evaevenro=" & l_evaevenro
   l_sql = l_sql & " AND evacab.empleado=" & l_ternro
   l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
   rsOpen l_rs_pa, cn, l_sql, 0 

	
	if l_rs_pa.eof then%>
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
		<td><b>Objetivo</b></td>
		<td><b>Qu&eacute; necesito alcanzar</b></td>
		<td><b>Pasos a Tomar para alcanzar el Objetivo</b></td>
		<td align="center"><b>Fecha para Cumplir el Objetivo</b></td>
		<td align="center"><b>Recursos para cumplir el Objetivo</b></td>
		<td align="center"><b>Brindará apoyo para cumplir el Objetivo</b></td>
	</tr>
	<tr>
		<td valign="top" height="55" class="texto">&nbsp;--</td>
		<td valign="top" height="55" class="texto">&nbsp;--</td>
		<td valign="top" height="55" class="texto">&nbsp;--</td>
		<td valign="top" align="center">&nbsp;--</td>
		<td valign="top" height="55" class="texto">&nbsp;--</td>
		<td valign="top" height="55" class="texto">&nbsp;--</td>
	</tr>
	</table>
	<%else%>
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<%end if
	do until l_rs_pa.eof%>
		
		<%l_evatevdesabr = l_rs_pa("evatevdesabr")
		if l_evatevnro<>l_rs_pa("evatevnro") then
			l_evatevnro = l_rs_pa("evatevnro") %>
		<tr>
		<td nowrap colspan="3"> <b><%= l_evatevdesabr %></b> </td>
		</tr>
		<tr>
		<td><b>Objetivo</b></td>
		<td><b>Qu&eacute; necesito alcanzar</b></td>
		<td><b>Pasos a Tomar para alcanzar el Objetivo</b></td>
		<td align="center"><b>Fecha para Cumplir el Objetivo</b></td>
		<td align="center"><b>Recursos para cumplir el Objetivo</b></td>
		<td align="center"><b>Brindará apoyo para cumplir el Objetivo</b></td>
	</tr>
	<%end if%>
	<tr>
		<td valign="top" height="55" class="texto"><%=l_rs_pa("evaobjdext")%>&nbsp;</td>
		<td valign="top" height="55" class="texto">&nbsp;</td>
		<td valign="top" height="55" class="texto">&nbsp;</td>
		<td valign="top" align="center">&nbsp;</td>
		<td valign="top" height="55" class="texto">&nbsp;</td>
		<td valign="top" height="55" class="texto">&nbsp;</td>
	</tr>

<%
	l_rs_pa.MoveNext
loop
l_rs_pa.Close%>
</table>
<%end function%>
