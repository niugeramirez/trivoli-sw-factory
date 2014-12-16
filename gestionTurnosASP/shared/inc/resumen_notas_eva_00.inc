<% 
'14-12-2004-CCRossi- mostrar el evaluador y ordenar por evaluador
'-----------------------------------------------------------------------------
function resumen_notas
 Dim l_evatevdesabr	
 Dim l_rsl
 	Set l_rsl = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evanotadesc, evatndesabr, evatndesext,evaseccnota.orden, "
	l_sql = l_sql & "evatipevalua.evatevdesabr "
	l_sql = l_sql & "FROM evanotas "
	l_sql = l_sql & "INNER JOIN evaseccnota ON evaseccnota.evatnnro = evanotas.evatnnro "
	l_sql = l_sql & "INNER JOIN evatiponota ON evatiponota.evatnnro = evanotas.evatnnro "
	l_sql = l_sql & "INNER JOIN evadetevldor ON evanotas.evldrnro=evadetevldor.evldrnro "
	l_sql = l_sql & "INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "INNER JOIN evasecc ON evadetevldor.evaseccnro=evasecc.evaseccnro "
	l_sql = l_sql & "WHERE evaseccnota.evaseccnro = " & l_evaseccnro
	l_sql = l_sql & " AND   evadetevldor.evacabnro = " & l_evacabnro
    l_sql = l_sql & " ORDER BY evatevdesabr, evaseccnota.orden "
    rsOpen l_rsl, cn, l_sql, 0
	
	if not l_rsl.eof then
		l_evatevdesabr = l_rsl("evatevdesabr")
	end if%>
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
		<td nowrap colspan="2"><b><%= l_evatevdesabr %></b></td>	
	</tr>
	<tr>
		<td width="50%"><b>Tipo Nota</b></td>
		<td width="50%"><b>Nota</b></td>
	</tr>
	
	<%	l_linea = l_linea + 2
	if l_rsl.eof then%>	
	<tr>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	</tr>
	<%	l_linea = l_linea + 1
	end if
   do while not l_rsl.eof 
	if l_evatevdesabr <> l_rsl("evatevdesabr") then
		l_evatevdesabr = l_rsl("evatevdesabr")%>
	<tr>
		<td nowrap colspan="2"><b><%= l_evatevdesabr %></b></td>	
	</tr>
	<tr>
		<td width="50%"><b>Tipo Nota</b></td>
		<td width="50%"><b>Nota</b></td>
	</tr>
		
	<%
	end if%>
   <tr>
		<td valign=top height="50" class="texto"><%=l_rsl("evatndesabr")%>&nbsp;</td>
		<td valign=top height="50" class="texto"><%=l_rsl("evanotadesc")%>&nbsp;</td>
   </tr>
  <%	l_linea = l_linea + 1
  l_rsl.Movenext
  loop
  %>
  </table>
  <%
  l_rsl.Close
  end function
  %>
