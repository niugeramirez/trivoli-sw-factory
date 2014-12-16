<% 
function form_plan_accion
 Dim l_evatevdesabr	
 Dim l_evatevnro
 
 Dim l_rs_pa
 	Set l_rs_pa = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT Distinct evatipevalua.evatevnro, evatipevalua.evatevdesabr "
	l_sql = l_sql & "FROM evaoblieva INNER JOIN evatipevalua ON evaoblieva.evatevnro=evatipevalua.evatevnro "
	l_sql = l_sql & "WHERE evaoblieva.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
	rsOpen l_rs_pa, cn, l_sql, 0 
	 l_evatevnro=""
    do while not l_rs_pa.eof 
		l_evatevdesabr = l_rs_pa("evatevdesabr")
		
		
	    if l_evatevnro<>l_rs_pa("evatevnro") then
			l_evatevnro = l_rs_pa("evatevnro")
	    %>
		<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
		<tr>
			<td nowrap colspan="3"> <b><%= l_evatevdesabr %></b> </td>
		</tr>
		<%end if%>
	<tr>
		<td width="35%"><b>Aspecto a Mejorar</b></td>
		<td width="50%"><b>Plan de Acción</b></td>
		<td width="15%"><b>Fecha Revisión</b></td>
	</tr>
	<tr>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" align="center">&nbsp;</td>
	</tr>
	<tr>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" align="center">&nbsp;</td>
	</tr>
	<tr>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" align="center">&nbsp;</td>
	</tr>
	<tr>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" align="center">&nbsp;</td>
	</tr>
	<tr>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" height="55" class="texto">&nbsp;</td>
	<td valign="top" align="center">&nbsp;</td>
	</tr>
  </table>
<%
	l_rs_pa.MoveNext
loop

l_rs_pa.Close
end function
%>
