<%
function calcular_fechas(cantdias, fecha) 

 dim  l_dia 
 dim  l_esferiado
 dim l_rs1
 dim l_sql
 
 l_canthab  = 0
 l_cantferi = 0


 l_fechasta = DateAdd("d", (cantdias - 1), cDate(fecha))
 'response.write("fechadesde=")
 'response.write(fecha)
 'response.write("<br>")
 'response.write("l_fechasta")
 'response.write(l_fechasta)
 'response.write("<br>")
 
 for l_dia = cDate(fecha) to cDate(l_fechasta)
		'response.write("<br>")
		'response.write(l_dia)
 			
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT * FROM feriado "
			l_sql = l_sql & "WHERE feriado.ferifecha = " & cambiafecha(l_dia,"","")
			rsOpen l_rs, cn, l_sql, 0 
			if not l_rs.eof then
				if l_rs("tipferinro") = 1  then 'Pais 
					Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
					l_sql = "SELECT * FROM pais "
					l_sql = l_sql & "WHERE pais.paisdef = -1 " 
					rsOpen l_rs1, cn, l_sql, 0 
					if not l_rs1.eof then
						if l_rs("fericodext") = l_rs1("paisnro") then
							l_esferiado = true
						end if
					end if
					l_rs1.Close
				end if	
			end if	
			l_rs.Close
	
            if l_esferiado or (weekday(l_dia) = 1) or (weekday(l_dia) = 7) then
				l_cantferi = l_cantferi + 1.
			else
				l_canthab = l_canthab + 1.
			end if	
	l_dia = DateAdd("m", 1, cDate(l_dia))			
 next
 
end function
%>