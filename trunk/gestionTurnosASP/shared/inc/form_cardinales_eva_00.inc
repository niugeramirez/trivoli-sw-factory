<% function form_cardinales

 Dim l_rs4
 Dim l_rs1
 dim l_estrnro
 
 if trim(l_ternro)<>"" then 
	'buscar la gerencia -----------------------------------------------------------------
	Set l_rs4 = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT estrdabr, htetdesde, estructura.estrnro  "
	l_sql = l_sql & " FROM his_estructura "
	l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
	l_sql = l_sql & " WHERE his_estructura.ternro=" & l_ternro
	l_sql = l_sql & " AND   his_estructura.tenro = 6 " 
	l_sql = l_sql & " AND   his_estructura.htethasta IS NULL " 
	l_sql = l_sql & " ORDER BY his_estructura.htetdesde DESC " 
	rsOpen l_rs4, cn, l_sql, 0
	if not l_rs4.eof then	
		l_estrnro  = l_rs4("estrnro")
	else
		l_estrnro  = ""
	end if	
	l_rs4.Close
	set l_rs4=nothing
end if%>

<table border="1" cellpadding="0" cellspacing="0">
	<tr style="border-color :CadetBlue;">
		<td><b>Descripci&oacute;n</b></td>
		<td><b>Ponderaci&oacute;n por Gerencia</b></td>
	</tr>

	<% 	'response.write ("<script>alert('"&l_evaseccnro&"')</script>")

   	  Set l_rs4 = Server.CreateObject("ADODB.RecordSet")
  	   l_sql = "SELECT evafactor.evafacnro, "
	   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, "
	   l_sql = l_sql & " evatitulo.evatitdesabr, "
	   l_sql = l_sql & " evaseccfactor.orden "
	   l_sql = l_sql & " FROM evaseccfactor "
	   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaseccfactor.evafacnro "
	   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
	   l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
	   l_sql = l_sql & " ORDER BY evaseccfactor.orden "
	   rsOpen l_rs4, cn, l_sql, 0
	'	response.write ("<script>alert('"&l_sql&"')</script>")
		do while not l_rs4.eof %>
		<tr>
		<td valign=middle><%=l_rs4("evafacdesabr")%></td>	
		<%	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evadescomp.evadcdes "
			l_sql = l_sql & " FROM evadescomp "
			l_sql = l_sql & " WHERE evadescomp.evafacnro = " & l_rs4("evafacnro")
			if trim(l_estrnro) <>"" then
			    l_sql = l_sql & " AND evadescomp.estrnro = " & l_estrnro
			end if
			rsOpen l_rs1, cn, l_sql, 0
			if not l_rs1.eof then%>
		<td valign=middle><%=l_rs1("evadcdes")%></td>
			<%else%>	
		<td valign=middle>?</td>
			<%end if
			  l_rs1.Close
			  set l_rs1 = nothing%>	
		</tr>
		<%	l_rs4.MoveNext
		loop
		l_rs4.close
		set l_rs4=nothing%>
</table>
<%end function%>

