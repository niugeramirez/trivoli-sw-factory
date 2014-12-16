<%

function resumen_gralobj
 on error goto 0
 dim l_rs_rol 
 dim l_rsgral
 dim l_rs2
%>  

<table border="1" cellpadding="0" cellspacing="0">
<%

Set l_rsgral = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evagralobj.evldrnro, evatevdesabr, evagralobj.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr "
l_sql = l_sql & " FROM evagralobj "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evagralobj.evldrnro "
l_sql = l_sql & "   AND evadetevldor.evacabnro= " & l_evacabnro
l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro=evadetevldor.evatevnro "
l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro=evagralobj.evatrnro "
l_sql = l_sql & " ORDER BY evatevdesabr "
rsOpen l_rsgral, cn, l_sql, 0
do while not l_rsgral.eof 
	%>

	<tr height="20">
		<td colspan=3>
			<b>
			<%=l_rsgral("evatevdesabr")%>&nbsp;-&nbsp;Evaluaci&oacute;n General de Objetivos:</b>
			&nbsp;<%=l_rsgral("evatrvalor")%>-<%=l_rsgral("evatrdesabr")%></b>
		</td>
	</tr>
	<tr>
		<td width="17%">Descripci&oacute;n</th>
		<%if cformed=-1 then%>
		<td width="17%">Forma de Medici&oacute;n</th>
		<%end if%>
		<%'Buscar Roles
		Set l_rs_rol = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT DISTINCT evatevdesabr "
		l_sql = l_sql & " FROM evaluaobj "
		l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
		l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
		l_sql = l_sql & " WHERE evadetevldor.evacabnro= " & l_evacabnro
		l_sql = l_sql & "  AND  evadetevldor.evaseccnro <> " & l_evaseccnro
		l_sql = l_sql & "  AND (evadetevldor.evatevnro = " & cautoevaluador
		l_sql = l_sql & "   OR  evadetevldor.evatevnro = " & cevaluador & ")"
		l_sql = l_sql & " ORDER BY evatipevalua.evatevdesabr "
		rsOpen l_rs_rol, cn, l_sql, 0
		do while not l_rs_rol.eof%>
			<td><%=l_rs_rol("evatevdesabr")%></td>        
		<%l_rs_rol.MoveNext
		loop
		l_rs_rol.Close
		set l_rs_rol=nothing%>
	</tr>	

	<%	
	l_linea = l_linea + 2
    Dim l_rsl
   'Hacer select de objetivos unicamente
	Set l_rsl = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evaobjetivo.evaobjnro, evaobjetivo.evaobjdext,evaobjformed "
	l_sql = l_sql & " FROM evaobjetivo "
	l_sql = l_sql & " INNER JOIN evaluaobj    ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
	l_sql = l_sql & "   AND evadetevldor.evacabnro= " & l_evacabnro
	l_sql = l_sql & " WHERE evaluaobj.evldrnro    = " & l_rsgral("evldrnro")
	rsOpen l_rsl, cn, l_sql, 0
	do while not l_rsl.eof %>
		<tr height="10">
			<td align=left width=60% valign=middle><%=trim(l_rsl("evaobjdext"))%></td>
			<%if cformed=-1 then%> 
				<td align=center valign=middle>
				<%=trim(l_rsl("evaobjformed"))%>
				</td>
			<%end if%>
			<% 'Buscar resultados  
			Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evatevdesabr, evaluaobj.evldrnro, evaluaobj.evaobjnro, "
			l_sql = l_sql & " evaluaobj.evatrnro , evatrdesabr, evatrvalor"
			l_sql = l_sql & " FROM evaluaobj "
			l_sql = l_sql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro "
			l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
			l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro = evatipevalua.evatevnro "
			l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro=evaluaobj.evatrnro "
			l_sql = l_sql & " WHERE evaluaobj.evaobjnro = " & l_rsl("evaobjnro")
			l_sql = l_sql & "  AND  evadetevldor.evaseccnro <> " & l_evaseccnro
			l_sql = l_sql & "  AND (evadetevldor.evatevnro = " & cautoevaluador
			l_sql = l_sql & "   OR  evadetevldor.evatevnro = " & cevaluador & ")"
			l_sql = l_sql & " ORDER BY evatipevalua.evatevdesabr "
			rsOpen l_rs2, cn, l_sql, 0
			'Response.Write l_sql
			do while not l_rs2.eof%>
				    <td width=20% nowrap valign=middle><%=l_rs2("evatrvalor")%>-<%=l_rs2("evatrdesabr")%></td>
		    <%l_rs2.MoveNext
		    loop
		    l_rs2.Close
		    set l_rs2=nothing%>
		</tr>
	<%
	l_rsl.MoveNext
	loop
	l_rsl.Close
	
l_rsgral.moveNext
loop
l_rsgral.Close
%>
	   
   
</table>

<%
end function
%>

