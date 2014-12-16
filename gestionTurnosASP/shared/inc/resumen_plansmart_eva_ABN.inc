<% 
function resumen_plansmartABN
 
dim l_evatevnro

 Dim l_rs_pa
 	Set l_rs_pa = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaplan.aspectomejorar, evaplan.planaccion, evaplan.planfecharev,  "
	l_sql = l_sql & " evatipevalua.evatevnro,evatipevalua.evatevdesabr, evaobliorden, "
	l_sql = l_sql & " evaobjetivo.evaobjdext, recursos, ayuda "
	l_sql = l_sql & " FROM evaplan "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evaplan.evldrnro=evadetevldor.evldrnro "
	if l_borrador=-1 and l_logeadoternro<>"" and l_super<>-1 then
	l_sql = l_sql & "		AND evadetevldor.evaluador = "  & l_logeadoternro
	end if
	l_sql = l_sql & " INNER JOIN evaobjetivo  ON evaplan.evaobjnro=evaobjetivo.evaobjnro "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
	if l_borrador<>-1 then
	l_sql = l_sql & "		AND  evatipevalua.evatevnro = " & cevaluador
	end if
	l_sql = l_sql & " INNER JOIN evaoblieva   ON evaoblieva.evatevnro=evadetevldor.evatevnro "
	l_sql = l_sql & "		 AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
	l_sql = l_sql & " WHERE evadetevldor.evacabnro  =" & l_evacabnro
	'l_sql = l_sql & " AND   evadetevldor.evaseccnro =" & l_evaseccnro
	l_sql = l_sql & " ORDER BY evaoblieva.evaobliorden "
	rsOpen l_rs_pa, cn, l_sql, 0 
	'response.write l_sql
	%>
	<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
	<tr>
		<%if l_super=-1 then%><td><b>Rol</b><%end if%></td>
		<td><b>Objetivo</b></td>
		<td><b>Qu&eacute; necesito alcanzar</b></td>
		<td><b>Pasos a Tomar para alcanzar el Objetivo</b></td>
		<td align="center"><b>Fecha para Cumplir el Objetivo</b></td>
		<td align="center"><b>Recursos para cumplir el Objetivo</b></td>
		<td align="center"><b>Brindará apoyo para cumplir el Objetivo</b></td>
	</tr>
	
	<%l_linea = l_linea + 2
	if l_rs_pa.eof then%>	
	<tr>
	<%if l_super=-1 then%><td><b>--></b><%end if%></td>
	<td valign="top" align="center">--</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	</tr>
	<%l_linea = l_linea + 1
	end if

do until l_rs_pa.eof%>	
	<tr>
	<%if l_super=-1 then%>
		<td>
			<%if l_evatevnro<>l_rs_pa("evatevnro") then%>
				<b><%=l_rs_pa("evatevdesabr")%></b>
			<%end if%>
			&nbsp;
		</td>
	<%end if%>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("evaobjdext")%></td>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("aspectomejorar")%>&nbsp;</td>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("planaccion")%>&nbsp;</td>
	<td valign="top" align="center"><%=l_rs_pa("planfecharev")%>&nbsp;</td>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("recursos")%>&nbsp;</td>
	<td valign="top" height="55" class="texto"><%=l_rs_pa("ayuda")%>&nbsp;</td>
	</tr>
<%  l_linea = l_linea + 1
	l_evatevnro=l_rs_pa("evatevnro")
	l_rs_pa.MoveNext
loop
  %>
  </table>
  <%

l_rs_pa.Close
end function
%>
