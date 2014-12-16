<% 
function resumen_vistos
 Dim l_evatevdesabr	
 Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT *  "
	l_sql = l_sql & " FROM  evavistos INNER JOIN evadetevldor ON evavistos.evldrnro=evadetevldor.evldrnro "
	l_sql = l_sql & "INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "WHERE evadetevldor.evacabnro =" & l_evacabnro & " and evadetevldor.evaseccnro =" & l_evaseccnro
   	rsOpen l_rsl, cn, l_sql, 0%>
   	
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	
	<%if l_rsl.eof then%>	
	<tr>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	</tr>
	<%	l_linea = l_linea + 1
	end if
   do until l_rsl.eof %>
	<tr>
		<td colspan="2">
			<b><%= l_rsl("evatevdesabr") %></b>
		</td>
	</tr>
	<tr>
		<td width="20%"><b>Fecha</b></td>
		<td width="80%"><b>Observaci&oacute;n</b></td>
	</tr>
   <tr>
		<td valign=top height="30" ><%=l_rsl("visfecha")%></td>
		<td valign=top height="30" class="texto"><%=l_rsl("visdesc")%></td>
   </tr>
  <%	l_linea = l_linea + 3
  l_rsl.Movenext
  loop
  %>
  </table>
  <%
  l_rsl.Close
  end function
  %>
