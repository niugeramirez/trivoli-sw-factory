<% 
function form_grafico
' de uso local  
  Dim i
  Dim j
%>  

<table width="50" align="center" border="1" cellpadding="0" cellspacing="0">

<%
i = 0 
do while i <= 100%>
	<tr height="2">
		<%if i = 0 then%>
			<td rowspan="21" width="10" class="grafico"><B>Potencial</B></td>
		<%end if %>
		<%if i = 0 then%>
			<td rowspan="7" width="10" class="grafico"><B>Alto</B></td>
		<%end if %>
		<%if i = 35 then%>
			<td rowspan="7" width="10" class="grafico"><B>Medio</B></td>
		<%end if %>
		<%if i = 70 then%>
			<td rowspan="7" width="10" class="grafico"><B>Bajo</B></td>
		<%end if %>
		
		<td width=2 align=center class="grafico"><%=(100 - i)%></td>
		<%
		  j = 0 
		  do while j <= 100%>
					<td width="2" class="grilla">&nbsp;</td>
		<%j = j + 5
		loop%>
	</tr>
	<%i = i + 5
loop

j = 0
%>
<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
    <%do while j <= 100%>
	   <td class="grafico"><%=j%></td>
	 <%j = j + 5
  loop%>
</tr>
</table>
<table width="50" align="center" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="7" class="grafico" align="center"><b>Bajo</b></td>
	<td colspan="7" class="grafico" align="center"><b>Medio</b></td>
	<td colspan="7" class="grafico" align="center"><b>Alto</b></td>
</tr>

<tr>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td class="grafico">&nbsp;</td>
	<td colspan="21" class="grafico" align="center"><b>Desempeño</b></td>
</tr>
</table>
<%
end function
%>