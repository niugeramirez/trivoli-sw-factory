<%
'  ==============================================================================
'  Para ABN
'  ==============================================================================

function resumen_generalABN
 
 dim cols
 cols=4
 
 if l_borrador=-1 then
	cols=cols+1
 end if
%>  
 
 
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
    <td width="17%">Rol</th>
    <td width="17%">Comentario</th>
    <td width="17%">Fecha</th>
</tr>	

<%	l_linea = l_linea + 1
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT puntaje, puntajemanual, puntajeevldrnro, visdesc, visfecha, evatevdesabr "
   l_sql = l_sql & " FROM evavistos "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evavistos.evldrnro "
   l_sql = l_sql & "	    AND evadetevldor.evacabnro =" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evacab  ON evacab.evacabnro = evadetevldor.evacabnro "
   l_sql = l_sql & "	    AND evacab.evacabnro =" & l_evacabnro
   l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro=evadetevldor.evatevnro "
   
   rsOpen l_rsl, cn, l_sql, 0
   'response.write(l_sql)

  if not l_rsl.eof then%>	
	<tr>
	<td valign="top" colspan="<%=cols%>" align="left"><b>Puntaje Final:</b>&nbsp;<%=l_rsl("puntajemanual")%></td>
	</tr>
  <% else %>	
	<tr>
	<td colspan="<%=cols%>">No cargo esta sección.</td>
	</tr>
  <%
  end if
  do while not l_rsl.eof%>
		<tr>
			<td><%=l_rsl("evatevdesabr")%></td>
			<td><%=l_rsl("visdesc")%>&nbsp;</td>
			<td><%=l_rsl("visfecha")%>&nbsp;</td>
		</tr>
	
	<%l_rsl.movenext
  loop
  l_linea = l_linea + 1
  l_rsl.close
  set l_rsl=nothing%>

</table>

<%
end function
%>

