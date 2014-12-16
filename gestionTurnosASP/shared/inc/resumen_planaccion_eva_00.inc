<% 
function resumen_plan_accion
 Dim l_evatevdesabr	

 Dim l_rs_pa
 	Set l_rs_pa = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaplan.aspectomejorar, evaplan.planaccion, evaplan.planfecharev,  "
	l_sql = l_sql & "evatipevalua.evatevnro, evatipevalua.evatevdesabr "
	l_sql = l_sql & "FROM evaplan INNER JOIN evadetevldor ON evaplan.evldrnro=evadetevldor.evldrnro "
	l_sql = l_sql & "INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "WHERE evadetevldor.evacabnro =" & l_evacabnro&" and evadetevldor.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
	rsOpen l_rs_pa, cn, l_sql, 0 

	if not l_rs_pa.eof then
		l_evatevdesabr = l_rs_pa("evatevdesabr")
	end if	%>
	
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
		<td nowrap colspan="3"> <b><%= l_evatevdesabr %></b> </td>
	</tr>
	<tr>
		<td><b>Aspecto a Mejorar</b></td>
		<td><b>Plan de Acción</b></td>
		<td align="center"><b>Fecha Revisión</b></td>
	</tr>
	<%l_linea = l_linea + 2

do until l_rs_pa.eof
%>	
	<tr>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("aspectomejorar")%></td>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("planaccion")%></td>
	<td valign="top" align="center"><%=l_rs_pa("planfecharev")%></td>
	</tr>
<%  l_linea = l_linea + 1
	l_rs_pa.MoveNext
loop
  %>
  </table>
  <%

l_rs_pa.Close
end function
%>
