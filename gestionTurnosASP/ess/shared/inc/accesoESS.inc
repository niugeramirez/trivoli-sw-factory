<% 

Dim l_ess_empleg
Dim l_ess_ternro
Dim l_ess_correcto
Dim l_ess_reportaa

l_ess_correcto = true
l_ess_empleg   = 0
l_ess_ternro   = 0
l_ess_reportaa = 0

if Session("empleg") = "" then
   l_ess_empleg = 0
	
   response.write "Usuario No Autorizado"
   response.end
else
   Set ess_rs = Server.CreateObject("ADODB.RecordSet")

   if request.QueryString("empleg") = "" then
      l_ess_empleg = Session("empleg")
	  l_ess_reportaa = 0

   else
      Dim ess_rs
	  Dim ess_sql

	  ess_sql = "SELECT empreporta,ternro FROM empleado WHERE empleg=" & request.QueryString("empleg")

	  rsOpen ess_rs, cn, ess_sql, 0


	  if ess_rs.eof then
	         l_ess_correcto = false
	  else

	     'if isNull(ess_rs("empreporta")) then
            '		l_ess_correcto = false
	     'else
            		l_ess_correcto = true
			l_ess_reportaa = ess_rs("empreporta")
			l_ess_ternro   = ess_rs("ternro")
	'     end if
	  end if
	  
	  ess_rs.close	  
	  
    end if
	
	if l_ess_correcto then
	
		  ess_sql = "SELECT ternro FROM empleado WHERE empleg=" & Session("empleg")
	
		  rsOpen ess_rs, cn, ess_sql, 0

		  if not ess_rs.eof then
		  
		     if CLng(l_ess_reportaa) = 0 then
	            l_ess_ternro = ess_rs("ternro")
                l_ess_correcto = true				
			 else
			    if Clng(l_ess_reportaa) = Clng(ess_rs("ternro")) then
	               l_ess_empleg = request.QueryString("empleg")
	               l_ess_correcto = true
				else
	               l_ess_correcto = false
				end if
			 end if
		  else
             l_ess_correcto = false
		  end if
		  
		  ess_rs.close	  
	end if
	
	set ess_rs = nothing
	
	if not l_ess_correcto then
		
		l_ess_empleg = 0
		response.write "Usuario No Autorizado"
		response.end
	end if

end if
%>