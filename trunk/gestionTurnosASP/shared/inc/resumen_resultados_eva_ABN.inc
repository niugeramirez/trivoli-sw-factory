<%
function resumen_resultadosABN

dim l_evatevnro
 
l_linea = l_linea + 1
 dim cols
 cols=5
 
 if l_borrador=-1 then
	cols=cols+1
	if l_super=-1 then
		cols=cols+1
	end if	
 end if
%>  

<table border="1" cellpadding="0" cellspacing="0">

<tr style="border-color :CadetBlue;">
	<%if l_super=-1 then%><td><b>Rol</b><%end if%></td>
	<td colspan=2><b>Factor</b></td>
	<td><b>Puntaje</b></td>
	<td><b>Descripci&oacute;n</b></td>
    <%if l_borrador=-1 and l_logeadoternro<>"" and l_super<>-1 then%>
    <td width="17%"><b>Observaci&oacute;n</b></th>
    <%end if%>

</tr>

<%
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT  evaresultado.evafacnro, evaresultado.evatrnro, evaresultado.evaresudesc, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, evatipresu.evatrdesabr, evatipresu.evatrvalor,"
   l_sql = l_sql & " evatitulo.evatitdesabr, evaseccfactor.orden,  evatipevalua.evatevnro,evatipevalua.evatevdesabr "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evaresultado.evldrnro=evadetevldor.evldrnro "
   if l_borrador=-1 and l_logeadoternro<>"" and l_super<>-1 then
   l_sql = l_sql & "		AND evadetevldor.evaluador = "  & l_logeadoternro
   end if
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   if l_borrador<>-1 then
   l_sql = l_sql & "		AND evatipevalua.evatevnro = " & cevaluador
   end if
   l_sql = l_sql & " INNER JOIN evatipresu ON evaresultado.evatrnro=evatipresu.evatrnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evadetevldor.evacabnro    = " & l_evacabnro
   if l_super=-1 then
   l_sql = l_sql & " ORDER BY evatipevalua.evatevnro "
   else
   l_sql = l_sql & " ORDER BY evaseccfactor.orden "
   end if
   rsOpen l_rsl, cn, l_sql, 0

   if l_rsl.eof then%>	
	<tr>
		<td colspan="<%=cols%>"><b>No hay Competencias EVALUADAS.</b></td>
	</tr>
	<%l_linea = l_linea + 1
  end if
  l_evatevnro=""
  do while not l_rsl.eof %>
  
		 <tr>
			<%if l_super=-1 then%>
				<td>
					<%if l_evatevnro<>l_rsl("evatevnro") then%>
						<b><%=l_rsl("evatevdesabr")%></b>
					<%end if%>
					&nbsp;
				</td>
			 <%end if%>
			 <td valign="top"><%=l_rsl("evafacdesabr")%>&nbsp;</td>	
			 <td valign="top" class="texto"><%=l_rsl("evafacdesext")%>&nbsp;</td>	
			 <td ><%= l_rsl("evatrvalor") %>&nbsp;</td>
			 <td ><%= l_rsl("evatrdesabr") %>&nbsp;</td>
			 <%if l_borrador=-1 and l_logeadoternro<>"" and l_super<>-1then%>
			 <td ><%= l_rsl("evaresudesc") %>&nbsp;</td>
			 <%end if%>
		</tr>
	<%
	l_linea = l_linea + 1
	l_evatevnro=l_rsl("evatevnro")
	l_rsl.MoveNext
loop %>
  </table>
  
<%l_rsl.Close
set l_rsl=nothing

if l_tieneobj=0 then
Set l_rsl = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT puntaje, puntajemanual "
l_sql = l_sql & " FROM evacab "
l_sql = l_sql & " WHERE evacabnro=" & l_evacabnro
rsOpen l_rsl, cn, l_sql, 0%>

 <tr>
	<td colspan=<%=cols%> valign="top" ><b>Puntuaci&oacute;n (0 a 5) = <%= l_rsl("puntaje") %>
	</b></td>	
 </tr>
</table>
<%
l_rsl.Close
set l_rsl=nothing

end if

end function
%>

