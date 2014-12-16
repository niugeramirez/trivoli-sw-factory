<% 
function form_notas
 Dim l_evatevdesabr	
 Dim l_evatevnro
 Dim l_rsl
 dim l_evatndesabr
 dim l_evatndesext
 	Set l_rsl = Server.CreateObject("ADODB.RecordSet")

	l_sql = "SELECT evatipevalua.evatevnro, evatipevalua.evatevdesabr, evatiponota.evatndesabr, evatiponota.evatndesext "
	l_sql = l_sql & " FROM evaoblieva   "
    l_sql = l_sql & " INNER JOIN evatipevalua ON evaoblieva.evatevnro=evatipevalua.evatevnro  "
	l_sql = l_sql & " INNER JOIN evaseccnota  ON evaseccnota.evaseccnro=evaoblieva.evaseccnro "
	l_sql = l_sql & " INNER JOIN evatiponota  ON evatiponota.evatnnro = evaseccnota.evatnnro"
	l_sql = l_sql & " WHERE evaoblieva.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
    rsOpen l_rsl, cn, l_sql, 0
    'response.write l_sql
    l_evatevnro=""
    do while not l_rsl.eof 
		l_evatevdesabr = l_rsl("evatevdesabr")
		
		
	    if l_evatevnro<>l_rsl("evatevnro") then
	    l_evatevnro = l_rsl("evatevnro")%>
		<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
		<tr>
			<td nowrap colspan="2"><b><%= l_evatevdesabr %></b></td>	
		</tr>
		<tr>
			<td width="50%"><b>Tipo Nota</b></td>
			<td width="50%"><b>Nota</b></td>
		</tr>
		<%end if
	
	   l_evatndesabr = l_rsl("evatndesabr") 
	   l_evatndesext = l_rsl("evatndesext")%>
   		<tr>
			<td valign=top height="50" class="texto">
				<%=l_evatndesext%>
			</td>
			<td valign=top height="50" class="texto">
				&nbsp;
			</td>
	   </tr>
	  <%l_rsl.Movenext
	  loop
  l_rsl.Close
	  %>
	 </table>
  <%
  end function
  %>
