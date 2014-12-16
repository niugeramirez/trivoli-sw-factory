<%
function resumen_resultados
 'HARCODED---------------------- !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'  l_evaseccnro = 4
'  l_evldrnro   = 1
%>  
<table border="1" cellpadding="0" cellspacing="0">
<tr>

<tr style="border-color :CadetBlue;">
	<td width="20%">Descripci&oacute;n</td>
	<td width="17%">Evaluador</td>
	<td width="12%">Resultado</td>
	<td width="51%">Observaci&oacute;n</td>
</tr>
	
	
<%
Dim l_rsl
   Set l_rsl = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT  evaresultado.evafacnro, evaresultado.evatrnro, evaresultado.evaresudesc, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, evatipresu.evatrdesabr, "
   l_sql = l_sql & " evatitulo.evatitdesabr, evaseccfactor.orden,  evatipevalua.evatevdesabr "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evaresultado.evldrnro=evadetevldor.evldrnro "
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " INNER JOIN evatipresu ON evaresultado.evatrnro=evatipresu.evatrnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evadetevldor.evacabnro    = " & l_evacabnro
   l_sql = l_sql & " ORDER BY evaseccfactor.orden "

   rsOpen l_rsl, cn, l_sql, 0

Dim textoant
Dim texto
textoant=""

   if not l_rsl.eof then
    	textoant =l_rsl("evafacdesabr")
		%>
		 <tr>
			 <td width="20%" valign="top">
			 <table border="0" cellpadding="0" cellspacing="0" height="150">
			 <tr>
			 	<td valign="top" height="40"><b><%= l_rsl("evafacdesabr") %></b></td>	
			 </tr>
			 <tr>
			 	<td valign="top" height="110" class="texto"><%= l_rsl("evafacdesext") %></td>	
			 </tr>
			 </table>
			 </td>
			 <td colspan="3">
			 <table border="1" cellpadding="0" cellspacing="0">
				<tr border=1>
					<td width="20%"> <b><%= l_rsl("evatevdesabr") %></b> </td>
					<td width="15%"><%= l_rsl("evatrdesabr") %></td>
					<td width="65%" height="50" class="texto"><%= l_rsl("evaresudesc") %>&nbsp;</td>
				</tr>
		<%
	end if 

   l_rsl.MoveNext

   do while not l_rsl.eof 
    	if textoant =l_rsl("evafacdesabr") then
		%>
				<tr border=1>
					<td width="20%"> <b><%= l_rsl("evatevdesabr") %></b> </td>
					<td width="15%"><%= l_rsl("evatrdesabr") %></td>
					<td width="65%" height="50" class="texto"><%= l_rsl("evaresudesc") %>&nbsp;</td>
				</tr>
		<%
			texto=""
		else
		%>
			 </table>
			</td>
		 </tr>
		 <tr>
			 <td width="20%" valign="top">
			 <table border="0" cellpadding="0" cellspacing="0" height="150">
			 <tr>
			 	<td valign="top" height="40"><b><%= l_rsl("evafacdesabr") %></b></td>	
			 </tr>
			 <tr>
			 	<td valign="top" height="110" class="texto"><%= l_rsl("evafacdesext") %></td>	
			 </tr>
			 </table>
			 </td>
			 <td colspan="3">
			 <table border="1" cellpadding="0" cellspacing="0">
				<tr border=1>
					<td width="20%"> <b><%= l_rsl("evatevdesabr") %></b> </td>
					<td width="15%"><%= l_rsl("evatrdesabr") %></td>
					<td width="65%" height="50" class="texto"><%= l_rsl("evaresudesc") %>&nbsp;</td>
				</tr>
		<%
			texto =	l_rsl("evafacdesabr")
			textoant=l_rsl("evafacdesabr")
		end if
%>
<%
	l_rsl.MoveNext
loop
  %>
			 </table>
			</td>
		 </tr>
  </table>
  <%

l_rsl.Close
end function
%>

