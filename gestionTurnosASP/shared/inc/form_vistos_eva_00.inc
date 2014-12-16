<% 
function form_vistos
 Dim l_evatevdesabr	
 Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
	l_sql = "SELECT evatipevalua.evatevnro, evatipevalua.evatevdesabr "
	l_sql = l_sql & "FROM evaoblieva INNER JOIN evatipevalua ON evaoblieva.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "WHERE evaoblieva.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
   	rsOpen l_rsl, cn, l_sql, 0
   %>
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
		<td width="15%"></td>
		<td width="30%"><b>Firma</b></td>
		<td width="15%"><b>Fecha</b></td>
		<td width="40%"><b>Observaci&oacute;n</b></td>
	</tr>
	<%

   do until l_rsl.eof %>
	<tr>
		<td>
			<b><%= l_rsl("evatevdesabr") %></b>
		</td>
		<td valign=top height="30" >&nbsp;</td>
		<td valign=top height="30" >&nbsp;</td>
		<td valign=top height="55" class="texto">&nbsp;</td>
   </tr>
  <%l_rsl.Movenext
  loop
  %>
  </table>
  <%
  l_rsl.Close
  end function
  %>
