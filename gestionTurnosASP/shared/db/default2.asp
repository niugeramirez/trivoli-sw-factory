<%
on error goto 0
 Dim cn
 Dim l_base
 Set cn = Server.CreateObject("ADODB.Connection")

 l_base = trim(request.Querystring("base"))

 Select Case l_base
        Case "1" 
			 cn.ConnectionString = "Provider=Ifxoledbc;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=rhproexp@rhsco;"
			 Session("base") = "1"
        Case "2" 
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=w2kbb;Initial Catalog=rhprox2"
			 'cn.ConnectionString = "Provider=sqloledb.1;server=w2kbb;database=rhpro;uid=sa;pwd="			 
			 Session("base") = "2"
        Case "3" 
			 'cn.ConnectionString = "Provider=IBMDADB2.1;Password=" & Request.Form("pass") & ";Persist Security Info=True;User ID=" & Request.Form("usr") & ";Data Source=rhprodb;"
			 'cn.ConnectionString = "Provider=IBMDADB2.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=rhpro;"	
			 cn.ConnectionString = "Provider=IBMDADB2.1;User ID=db2admin;Data Source=rhpro;Persist Security Info=False"
	         Session("base") = "3"
        Case "4" 
			 cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=false;User ID=" & Request.Querystring("usr") & ";Data Source=rhprox2;"
			 Session("base") = "4"
		Case "5" 
			 cn.ConnectionString = "Provider=Ifxoledbc;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=megatone@rhsco;"
			 'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=localhost;Initial Catalog=base0"
			 'cn.ConnectionString = "Provider=sqloledb.1;server=RHWEBBB\RHPRO;database=rhpro;uid=sa;pwd="
			 Session("base") = "5"	 
		Case "6" 
			 cn.ConnectionString = "Provider=MySqlProv;Password=" & Request.Querystring("pass") & ";User ID=" & Request.Querystring("usr") & ";Data Source=rhprox2;Location=rhdesa;"
			 Session("base") = "6"	 
        Case "7" 
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=localhost;Initial Catalog=base0"
			 Session("base") = "7"	 
        Case "8" 
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=w2kbb;Initial Catalog=PromoFilmNew"
			 Session("base") = "8"	 
        Case "9" 
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=w2kbb;Initial Catalog=PromoFilm"
			 Session("base") = "9" 
        Case "10" 
'			cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=w2kbb;Initial Catalog=PromoFilm"
			cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=w2kbb;Initial Catalog=citrusvil"
             Session("base") = "10"
 End Select

'---Menu---'
 ON ERROR resume next
 Dim l_menu
 Dim l_debug
 l_menu = trim(request.Querystring("menu"))
 l_debug = CInt(request.Querystring("debug"))
 cn.Open
 if err then
	if l_menu = "html" then
		if l_debug = -1 then
			%><script>parent.document.FormVar.desc.value = "<%= Err.Description %>";</script><%
		else
			%><script>parent.document.FormVar.desc.value = "Acceso no valido";</script><%
		end if
	else
		if l_debug = -1 then
			response.write "&acceso=" & Err.Description & "&"
		else
			response.write "&acceso=Acceso no valido&"
		end if
	end if
 else
	Dim l_rs, l_sql
	Dim l_iduser
	Dim l_pass
	Dim l_tiempo
	Dim l_fecha
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_iduser = Request.Querystring("usr")
	l_pass	 = Request.Querystring("pass")
	l_tiempo = now
	
	Session("UserName") = l_iduser
	Session("Password") = l_pass
	Session("Time") = l_tiempo
	l_fecha = date()
	l_hora 	= hour(l_fecha) & ":" & minute(l_fecha) & ":" & second(l_fecha)
	
	' Se ingresa en la base de datos, la fecha y hora de logueo.
	l_sql = "SELECT hlognro FROM hist_log_usr WHERE iduser = " & Session("UserName")
	rsOpen l_rs, cn, l_sql, 0
	if l_rs.eof then
		l_sql = 		"INSERT INTO hist_log_usr (iduser, hlogfecini, hloghoraini) "
		l_sql = l_sql & " VALUES (" & l_iduser & ", " & l_fecha & ", " & l_hora & ")"
	else
	
	end if
	
	
	response.write "&acceso=Valido&"
	if l_menu = "html" then
		%><script>
			parent.document.location = "../../lanzador/lanzador3.asp";
		</script><%
	end if

end if
cn.close
%>
