<%
Dim Debug
Dim SegBotones
Debug = true
SegBotones = true


function fsql_join(sql)
	dim fin
	dim inicio
	dim l_join
	dim tabla1
	dim joinon1
	dim sqltemp
	dim sqltemp2
	dim sqlnew
	dim corteEn

	if instr(ucase(sql),"INNER JOIN") <> 0 then
		do until instr(ucase(sql),"INNER JOIN") = 0
			replace sql,"  "," "
			inicio = instr(ucase(sql),"INNER JOIN")
			fin = instr(inicio,ucase(sql),"WHERE") - 1
			if fin = -1 then 
				fin = instr(inicio,ucase(sql),"GROUP BY") - 1
				if fin = -1 then 
					fin = instr(inicio,ucase(sql),"ORDER BY") - 1
					if fin = -1 then
						fin = len(sql)
						corteEn = "EOL"
					else
						corteEn = "ORDER BY"
					end if
				else
					corteEn = "GROUP BY"
				end if
			else
				corteEn = "WHERE"
			end if

			l_join = trim(mid(sql,inicio,fin - inicio))
			l_join = mid(sql,inicio,fin - inicio)
			tabla1 = mid(l_join,12,len(sql))
			joinon1 = mid(tabla1,instr(ucase(tabla1)," ON ") + 4,len(sql))
			tabla1 = left(tabla1,instr(tabla1," ")-1)
			
			sqltemp = left(sql,inicio - 1)
			sqltemp = sqltemp & " " & right(sql,len(sql)-fin)
			sqltemp2 = left(sqltemp,instr(ucase(sqltemp),"FROM") + 4) & tabla1 & ","
			sqltemp2 = sqltemp2 & right(sqltemp,len(sqltemp) - instr(ucase(sqltemp),"FROM") - 4)
			if corteEn = "WHERE" then
				sqlnew = left(sqltemp2,instr(ucase(sqltemp2),corteEn) + 5) & "(" & joinon1 & ") AND "
			else
				sqlnew = sqltemp2 & " WHERE (" & joinon1 & ") "
			end if
			if corteEn <> "EOL" then
				sqlnew = sqlnew & right(sqltemp2,len(sqltemp2) - instr(ucase(sqltemp2),corteEn) - 4)
			end if
			sql = sqlnew
		loop
		fsql_join = sqlnew
	else
		fsql_join = sql
	end if
end function

'Abre un recordset con el Tipo de Cursor dado
'Esto permite abrir un cursor con desplazamiento en todo sentido
sub rsOpenCursor(ByRef rs, ByVal cn, ByVal sql, ByVal nivel, TipoCursor)
	rs.CursorType = TipoCursor
	call rsOpen(rs,cn,sql,nivel)	
end sub

Sub rsOpen(ByRef rs, ByVal cn, ByVal sql, ByVal nivel)
	Dim Buffer
	on error resume next
	if l_base = "4" then sql=fsql_join(sql)
	RS.Open sql, cn
	if err then
		Buffer = "<script>"
		if debug then 
			Buffer = Buffer & "alert(" & chr(34) & "Debug: TRUE\n"
			Buffer = Buffer & "Archivo: " & Request.ServerVariables("SCRIPT_NAME") & "\n"
			Buffer = Buffer & "Numero Error: " & err.number & "\n"
			Buffer = Buffer & "Descripcion: " & err.description & "\n"
			Buffer = Buffer & "SQL: " & sql & chr(34) & ");"
			Buffer = Buffer & "prompt('SQL String:'," & chr(34) & sql & chr(34) & ");"
		else
			Buffer = Buffer & "alert('" & err.description & "');"
		end if
		if nivel = 1 then
			Buffer = Buffer & "window.close();"
		else
			if nivel = 2 then
				Buffer = Buffer & "window.opener.close();window.close();"
			end if
		end if
		Buffer = Buffer & "</script>"
		response.write Buffer
		response.end
	end if
End Sub

function killChars(strWords)
 dim badChars 
 dim newChars
 dim i
   badChars = array("drop", ";", "--", chr(10), chr(13), chr(34), "xp_") 
   newChars = strWords
   for i = 0 to uBound(badChars) 
     newChars = replace(newChars, badChars(i), " ") 
   next
   killChars = newChars
end function

Sub cmExecute(ByRef cm, ByVal sql, ByVal nivel)
	Dim Buffer
    sql = killChars(sql)
'	response.write "<script>alert('" & sql & "')</script>"
	on error resume next
	cm.commandText = sql
	cm.Execute
	if err then
		Buffer = "<script>"
		if debug then 
			Buffer = Buffer & "alert(" & chr(34) & "Debug: TRUE\n"
			Buffer = Buffer & "Archivo: " & Request.ServerVariables("SCRIPT_NAME") & "\n"
			Buffer = Buffer & "Numero Error: " & err.number & "\n"
			Buffer = Buffer & "Descripcion: " & err.description & "\n"
			Buffer = Buffer & "SQL: " & sql & chr(34) & ");"
			Buffer = Buffer & "prompt('SQL String:',"&chr(34)&sql&chr(34)&");"
		else
			Buffer = Buffer & "alert('" & err.description & "');"
		end if
		if nivel = 1 then
			Buffer = Buffer & "window.close();"
		else
			if nivel = 2 then
				Buffer = Buffer & "window.opener.close();window.close();"
			end if
		end if
		Buffer = Buffer & "</script>"
		response.write Buffer
		response.end
	end if
End Sub

Dim cn 
Set cn = Server.CreateObject("ADODB.Connection")
Dim l_base
l_base = 2'Session("base")
'l_base = Session("base_z")

'response.write "<script>window.alert(' " & l_base & " ');</script>"

Select Case l_base
        Case "1" 
  		   cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=bb-omh-sis012;Initial Catalog=ticketMM"				
        Case "2" 

			cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=127.0.0.1;Initial Catalog=wi791471_swimming"
			'cn.ConnectionString = "Driver={SQL Server Native Client 10.0};Server=localhost;Database=wi521535_turnos;User ID=" & Session("username") & ";Password=" & Session("password")
			'cn.ConnectionString = "Driver={SQL Server Native Client 10.0};Server=localhost;Database=wi521535_turnos;Trusted_Connection=yes;"
			'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=bb-omh-pc009;Initial Catalog=ticketprueba"
			'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=bb-omh-tkt002;Initial Catalog=ticket"
			'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=True;User ID=" & Session("username") & ";Data Source=RHWEBBB\RHPRO;Initial Catalog=rhprox2"
        Case "3"'dx 
 		   'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=bb-omh-tkt003;Initial Catalog=ticket"
        Case "4"'et 
 			'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Session("password") & ";Persist Security Info=true;User ID=" & Session("username") & ";Data Source=bb-omh-tkt003;Initial Catalog=ticket"
			
End Select
on error resume next

'response.write "<script>window.alert(' " & cn.ConnectionString & " ');</script>"

cn.Open
if l_base = "3" then
		Dim rs97
		Dim sql97
		Set rs97 = Server.CreateObject("ADODB.RecordSet")
		sql97 = "Set schema db2admin"
		rs97.Open sql97, Cn
end if
if err then
		response.write "<script>window.alert('Problemas con la conexión. Por favor salga del Sistema y vuelva a ingresar.');window.close;</script>"
		response.write "<script>window.alert("""& err.Description &""");window.close;</script>"
	response.end
else
	Session("Time") = now
end if

' Muestra la entrada en el menú si tiene la opción habilitada
sub MostrarBoton (Style, Pagina, Nombre)
  dim Actual
  dim perfil
  dim username
  dim btnaccess
  dim myRs
  dim sql
  dim autorizado

  autorizado = Habilitado(Nombre)
  if (not autorizado) then  
    'response.write("<a class=sidebtnDSB href=""" & "#" & """>" & Nombre & "</a>")
  else
    response.write("<a class=" & Style & " href=""" & Pagina & """>" & Nombre & "</a>")
  end if

end sub

' Muestra la entrada en el menú como disponible si tiene la opción habilitada
sub MostrarOpcionMenu (Style, Pagina, Nombre, fnClick, fnOut)
  dim autorizado
  autorizado = Habilitado(Nombre)
  if (not autorizado) then  
    response.write("<a class=opcionbtnno href=""" & "#" & """>" & Nombre & "</a>")
  else
    response.write("<a class=" & Style & " href=""" & "#" & """ onclick=""" & fnClick & """ onmouseout=""" & fnOut & """ >" & Nombre & "</a>")
  end if

end sub

function Habilitado(Nombre)
  dim Actual
  dim perfil
  dim username
  dim btnaccess
  dim myRs
  dim sql
  dim autorizado
  dim i
  dim arreglo

  if SegBotones then
	  Set myRs = Server.CreateObject("ADODB.RecordSet")
	  
	  username = UCase(Session("Username"))
	  if not username = "SUPER" then
		sql = "SELECT perfnom FROM user_per"
		sql = sql & " inner join perf_usr on perf_usr.perfnro = user_per.perfnro"
		sql = sql & " where upper(iduser) = '" & username & "'"
		myRs.Maxrecords = 1
		rsOpen myRs, cn, sql, 0
		perfil = myRs("perfnom")
		myRs.Close
	
	    Actual = Request.ServerVariables("SCRIPT_NAME")
	    Actual = ucase(right(actual, len(actual) - instrrev(Actual,"/")))
	
	    sql = "SELECT btnaccess FROM menubtn"
	    sql = sql & " where btnpagina = '" & actual & "' and btnnombre = '" & ucase(nombre) & "'"
	    myRs.Maxrecords = 1
	    rsOpen myRs, cn, sql, 0
	    if myRs.eof then
	      autorizado = true
	    else
	
	      if (myRs("btnaccess") = perfil) or (myRs("btnaccess") = "*") then
		    autorizado = true
		  else
		    arreglo = split(myRs("btnaccess"),";")
			
			autorizado = false

			for i=0 to UBound(arreglo)
			   if UCase(arreglo(i)) = UCase(perfil) then
			      autorizado = true
			   end if
			next

		  end if
	    end if
	    myRs.Close
	  else
	    autorizado = true
	  end if
  else
    autorizado = true  
  end if
  Habilitado = autorizado
end function

' Muestra la entrada en el menú si tiene la opción habilitada
sub MostrarOpcion (Pagina,Nombre)
  dim Actual
  dim perfil
  dim username
  dim btnaccess
  dim myRs
  dim sql
  dim autorizado

  autorizado = Habilitado(Nombre)
'  autorizado = true
  if (not autorizado) then  
    response.write("")
  else
    response.write(Pagina)
  end if
end sub

%>
