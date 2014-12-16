<%
 Dim cn
 Dim l_base
 Set cn = Server.CreateObject("ADODB.Connection")
 l_base = trim(request.Querystring("base"))
 Select Case l_base
        Case "1" 
			 cn.ConnectionString = "Provider=Ifxoledbc;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=rhproexp@rhsco;"
			 Session("base") = "1"
        Case "2" 
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sa;Data Source=bb-omh-tkt001;Initial Catalog=Ticket"
			 'cn.ConnectionString = "Provider=sqloledb.1;server=RHWEBBB\RHPRO;database=rhpro;uid=sa;pwd="			 
			 Session("base") = "2"
        Case "3" 
			 'cn.ConnectionString = "Provider=IBMDADB2.1;Password=" & Request.Form("pass") & ";Persist Security Info=True;User ID=" & Request.Form("usr") & ";Data Source=rhprodb;"
			 cn.ConnectionString = "Provider=IBMDADB2.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=azul1;"	
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
			 cn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Request.Querystring("pass") & ";Persist Security Info=True;User ID=" & Request.Querystring("usr") & ";Data Source=localhost;Initial Catalog=Plasnavi"
			 Session("base") = "9" 



 End Select
 ON ERROR resume next
 cn.Open
if err then
	response.write "&acceso=NoValido&"
else
	Dim rs, sql
	Session("UserName") = Request.Querystring("usr")
	Session("Password") = Request.Querystring("pass")
	Session("Time") = now
	response.write "&acceso=Valido&"
end if
cn.close
%>
