<% function form_objetivos %>  
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
<tr>
    <td width="17%">Descripci&oacute;n</th>
    <td width="17%">Forma de Medici&oacute;n</th>
    <td width="17%">Evaluador</td>
	<td width="12%">Resultado</td>
</tr>	
<%
   Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaobjetivo.evaobjnro,evaperfijo, evapernroeva, evaobjdext,evaobjformed, "
   l_sql = l_sql & " evaluaobj.evldrnro, "
   l_sql = l_sql & " evatipevalua.evatevdesabr, evatrdesabr "
   l_sql = l_sql & " FROM evaobjetivo "
   l_sql = l_sql & " INNER JOIN evaluaobj ON evaluaobj.evaobjnro = evaobjetivo.evaobjnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
   l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaluaobj.evatrnro "
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " AND evadetevldor.evacabnro=" & l_evacabnro
   rsOpen l_rsl, cn, l_sql, 0

 do while not l_rsl.eof %>
 <tr>
 	<td valign="top" ><b><%= l_rsl("evaobjdext") %></b></td>	
 	<td valign="top" class="texto"><%= l_rsl("evaobjformed") %></td>	
	<td width="20%"> <b><%= l_rsl("evatevdesabr") %></b> </td>
	<td width="15%">&nbsp;</td>
</tr>
<%	l_rsl.MoveNext
loop   %>
			 </table>
			</td>
		 </tr>
  </table>
  <%

l_rsl.Close

end function
%>

