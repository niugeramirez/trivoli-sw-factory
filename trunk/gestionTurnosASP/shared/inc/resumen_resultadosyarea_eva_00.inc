<% 
'--------------------------------------------------------------------------
'Archivo       : resumen_resultadosyarea_eva_00.inc
'Descripcion   : resumen de ... 
'Creacion      : 04 - 01- 2004  
'Autor         : Leticia Amadio.
'Modificacion  : 
'--------------------------------------------------------------------------


function resumen_resultadosyarea
 
	l_linea = l_linea + 1  ' VERRRRRRRRRRRRRRRRR 
 
 Dim  l_evatitdesabr
 Dim  l_evatevdesabr
 Dim l_rsra
 Dim l_rs1
 
 Dim l_areatrnro   
 Dim l_evaareadesc 
 Dim l_interpretaciones
 Dim l_observables 
 
 Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
%>  


<%
	 'BUSCAR evaresultados     
   Set l_rsra = Server.CreateObject("ADODB.RecordSet")
   l_sql = "SELECT evaresultado.evldrnro, evaresultado.evafacnro, "
   l_sql = l_sql & " evaresultado.evatrnro, evaresultado.evaresudesc, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
   l_sql = l_sql & " evatitulo.evatitnro, evatitulo.evatitdesabr, "
   l_sql = l_sql & " evaseccfactor.orden , evatevdesabr"
   l_sql = l_sql & " , evadetevldor.evldrnro "
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evaresultado.evldrnro=evadetevldor.evldrnro "
   l_sql = l_sql & " INNER JOIN evatipevalua ON evadetevldor.evatevnro=evatipevalua.evatevnro "
   l_sql = l_sql & " INNER JOIN evatipresu ON evaresultado.evatrnro=evatipresu.evatrnro "
   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
   l_sql = l_sql & " AND   evadetevldor.evacabnro    = " & l_evacabnro
   l_sql = l_sql & " ORDER BY evatevdesabr "
   
   l_evatitdesabr=""
   rsOpen l_rsra, cn, l_sql, 0
 %>
   
   <table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="#FAF0E6">
<%
	l_evatevdesabr = ""
  'if not l_rsra.eof then
  	'l_evatedesabr= l_rsra("evatevdesabr")
  'end if
	
   Do While not l_rsra.eof 
   		
	if l_evatevdesabr <> l_rsra("evatevdesabr") then
		
		l_evatevdesabr = l_rsra("evatevdesabr")
		%>
			<tr><td colspan="5"> <b><%=l_rsra("evatevdesabr")%></b></td></tr>
		<%
		
		if trim(l_evatitdesabr) <> trim(l_rsra("evatitdesabr")) then
		%>
			<tr>
				<td align="right" width="30%"><b><%=l_rsra("evatitdesabr")%></b></td>
		<%  
					'buscar evaarea  ----------------------------
				    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT evatrnro, evaareadesc  "
					l_sql = l_sql & " FROM evaarea "
					l_sql = l_sql & " WHERE evaarea.evatitnro  = " & l_rsra("evatitnro")
					l_sql = l_sql & " AND   evaarea.evldrnro  = " & l_rsra("evldrnro") 'l_evldrnro
					rsOpen l_rs1, cn, l_sql, 0
					if not l_rs1.eof then
						l_areatrnro=l_rs1("evatrnro")
						l_evaareadesc=l_rs1("evaareadesc")
					end if
					l_rs1.close
					set l_rs1=nothing
					
			    	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT DISTINCT evaresu.evatrnro,  "
					l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
					l_sql = l_sql & " FROM evaresu "
					l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
					l_sql = l_sql & " INNER JOIN evafactor  ON evafactor.evafacnro = evaresu.evafacnro "
					l_sql = l_sql & " INNER JOIN evatitulo  ON evatitulo.evatitnro = evafactor.evatitnro "
					l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
					l_sql = l_sql & " AND   evatitulo.evatitnro  = " & l_rsra("evatitnro")
					l_sql = l_sql & " AND evaresu.evatrnro= " & l_areatrnro
					l_sql = l_sql & " order by evatrvalor "
					rsOpen l_rs1, cn, l_sql, 0
				%>	
				<td align="left" width="20%"><%=l_rs1("evatrdesabr")%></td>
				<%
					l_rs1.Close
					set l_rs1 = nothing
				%>
				<td colspan="3"><%=trim(l_evaareadesc)%> &nbsp;</td>
			</tr>
			<tr>
				<td><b>Descripci&oacute;n</b></td>
				<td><b>Resultado</b></td>
				<td colspan="3"><b>Observaci&oacute;n</b></td>
			</tr>
			<%	
				l_evatitdesabr = l_rsra("evatitdesabr")
		end if
		%>
		
		<tr>
			<td><%=l_rsra("evafacdesabr")%></td>	
			<%
			  		'BUSCAR la descripcion de evaresu  ----------------------------
			    Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT evaresu.evatrnro, evaresu.evaresudes, "
				l_sql = l_sql & " evatipresu.evatrvalor, evatipresu.evatrdesabr "
				l_sql = l_sql & " FROM evaresu "
				l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresu.evatrnro "
				l_sql = l_sql & " WHERE evaresu.evaseccnro = " & l_evaseccnro
				l_sql = l_sql & " AND   evaresu.evafacnro  = " & l_rsra("evafacnro")
				l_sql = l_sql & " AND   evaresu.evatrnro = " & l_rsra("evatrnro")
				l_sql = l_sql & " order by evatrvalor "
				rsOpen l_rs1, cn, l_sql, 0
			%>
				<td nowrap> <%=l_rs1("evatrdesabr")  %> </td>
			<%
				l_rs1.Close 
				set l_rs1 = nothing 
	 		%>
			
			<td colspan="3"><%=trim(l_rsra("evaresudesc"))%> &nbsp;</td>
			
		</tr>
<%

		end if  	' l_evatedesabr
	%>
		<tr><td colspan="5">&nbsp;</td></tr>
	<%
		l_rsra.Movenext
		loop
		

		l_rsra.Close
		
	%>

</table>

<%
end function
%>

