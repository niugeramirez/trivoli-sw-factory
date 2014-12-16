<%
function resumen_objetivos
 on error goto 0 
%>  
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
    <td width="17%">Descripci&oacute;n</th>
    <td width="17%">Forma de Medici&oacute;n</th>
    <td width="17%">Evaluador</td>
	<td width="12%">Resultado</td>
</tr>	
<%	l_linea = l_linea + 1
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT distinct evaobjetivo.evaobjnro,evaobjdext,evaobjformed, "
   l_sql = l_sql & " evatipevalua.evatevdesabr, evatipevalua.evatevnro, evatrdesabr "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & "	AND evadetevldor.evacabnro=" & l_evacabnro
   l_sql = l_sql & "	AND evadetevldor.evldrnro = evaluaobj.evldrnro" 
   l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaluaobj.evatrnro "
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " ORDER BY evatipevalua.evatevnro"
   rsOpen l_rsl, cn, l_sql, 0
'response.write(l_sql)
if l_rsl.eof then%>	
	<tr>
	<td valign="top" align="center">--</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	<td valign="top" align="center">--&nbsp;</td>
	</tr>
<%	l_linea = l_linea + 1
 end if
 do while not l_rsl.eof %>
 <tr>
 	<td valign="top" ><b><%= l_rsl("evaobjdext") %></b></td>	
 	<td valign="top" class="texto"><%= l_rsl("evaobjformed") %></td>	
	<td width="20%"><%= l_rsl("evatevdesabr") %> </td>
	<td width="15%"><%= l_rsl("evatrdesabr") %></td>
</tr>
<%	l_linea = l_linea + 1
l_rsl.MoveNext
loop   %>
			 </table>
			</td>
		 </tr>
  </table>
  <%

l_rsl.Close

end function
%>

