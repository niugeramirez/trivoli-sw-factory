<% 
function form_totales
 Dim l_evatevdesabr	
 Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
	l_sql = "SELECT evatipevalua.evatevnro, evatipevalua.evatevdesabr "
	l_sql = l_sql & "FROM evaoblieva INNER JOIN evatipevalua ON evaoblieva.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "WHERE evaoblieva.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
   	rsOpen l_rsl, cn, l_sql, 0
%>	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
	<td align=center>&nbsp;</td> 
	<td align=center>Suma Total</td> 
    <td align=center>% Total</td>
	<td align=center>Calif. Final</td>
	</tr>
<%
   do until l_rsl.eof %>
	<tr>
	<td nowrap> <b><%= l_rsl("evatevdesabr") %></b> </td>
	<td nowrap align=center>&nbsp;</td>
	<td nowrap align=center>&nbsp;</td>
	<td nowrap align=center>&nbsp;</td>
	</tr>
  <%l_rsl.Movenext
  loop
%>
  </table>

<%
  l_rsl.Close
  end function
  %>
